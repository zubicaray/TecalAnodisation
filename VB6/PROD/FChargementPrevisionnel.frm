VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "picclp32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FChargementPrevisionnel 
   ClientHeight    =   13005
   ClientLeft      =   -75
   ClientTop       =   2085
   ClientWidth     =   28080
   Icon            =   "FChargementPrevisionnel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13005
   ScaleWidth      =   28080
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12795
      Index           =   0
      Left            =   0
      ScaleHeight     =   12795
      ScaleWidth      =   28080
      TabIndex        =   3
      Top             =   375
      Width           =   28080
      Begin VB.PictureBox PBDeplacementFenetre 
         Height          =   12675
         Index           =   1
         Left            =   -15
         ScaleHeight     =   12615
         ScaleWidth      =   28500
         TabIndex        =   4
         Top             =   0
         Width           =   28560
         Begin C1SizerLibCtl.C1Tab CTOnglets 
            Height          =   12255
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   28215
            _cx             =   49768
            _cy             =   21616
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
            Caption         =   "Chargement|Prévisionnel"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   1
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   2
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
            Picture(0)      =   "FChargementPrevisionnel.frx":014A
            Picture(1)      =   "FChargementPrevisionnel.frx":02A4
            Begin VB.PictureBox PBOnglets 
               Height          =   11715
               Index           =   1
               Left            =   28860
               ScaleHeight     =   11655
               ScaleWidth      =   28065
               TabIndex        =   41
               Top             =   495
               Width           =   28125
               Begin VB.CommandButton CBCalculerPrevisionnel 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Calculer"
                  DownPicture     =   "FChargementPrevisionnel.frx":0DEE
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   975
                  Left            =   8100
                  MaskColor       =   &H00FF00FF&
                  Picture         =   "FChargementPrevisionnel.frx":1538
                  Style           =   1  'Graphical
                  TabIndex        =   124
                  Top             =   480
                  UseMaskColor    =   -1  'True
                  Width           =   1935
               End
               Begin VB.PictureBox PBCriteresRecherche 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808080&
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
                  Height          =   1515
                  Index           =   1
                  Left            =   180
                  ScaleHeight     =   1485
                  ScaleWidth      =   7425
                  TabIndex        =   105
                  Top             =   180
                  Width           =   7455
                  Begin VB.CommandButton CBRaz 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "RAZ"
                     BeginProperty Font 
                        Name            =   "Small Fonts"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   555
                     Index           =   1
                     Left            =   5460
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":1C82
                     Style           =   1  'Graphical
                     TabIndex        =   110
                     ToolTipText     =   " Annule tris et recherches "
                     Top             =   780
                     UseMaskColor    =   -1  'True
                     Width           =   1755
                  End
                  Begin VB.CommandButton CBLancerRecherche 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "RECHERCHER"
                     BeginProperty Font 
                        Name            =   "Small Fonts"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   555
                     Index           =   1
                     Left            =   5460
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":1E74
                     Style           =   1  'Graphical
                     TabIndex        =   109
                     ToolTipText     =   " Lancer une recherche "
                     Top             =   120
                     UseMaskColor    =   -1  'True
                     Width           =   1755
                  End
                  Begin VB.TextBox TBCommencantPar 
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
                     Height          =   315
                     Index           =   1
                     Left            =   2100
                     TabIndex        =   108
                     Top             =   600
                     Width           =   3135
                  End
                  Begin VB.TextBox TBContenant 
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
                     Height          =   315
                     Index           =   1
                     Left            =   2100
                     TabIndex        =   107
                     Top             =   1020
                     Width           =   3135
                  End
                  Begin VB.ComboBox CBRechercherPar 
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
                     Height          =   360
                     Index           =   1
                     ItemData        =   "FChargementPrevisionnel.frx":21B6
                     Left            =   2100
                     List            =   "FChargementPrevisionnel.frx":21C3
                     Style           =   2  'Dropdown List
                     TabIndex        =   106
                     Top             =   120
                     Width           =   3135
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rechercher par"
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
                     Height          =   255
                     Index           =   30
                     Left            =   0
                     TabIndex        =   113
                     Top             =   180
                     Width           =   1935
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Commençant par"
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
                     Height          =   255
                     Index           =   29
                     Left            =   180
                     TabIndex        =   112
                     Top             =   660
                     Width           =   1755
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Contenant"
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
                     Height          =   255
                     Index           =   27
                     Left            =   840
                     TabIndex        =   111
                     Top             =   1080
                     Width           =   1095
                  End
               End
               Begin VB.PictureBox PBPrevisionnel 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00C0C0C0&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   11295
                  Left            =   10500
                  ScaleHeight     =   11295
                  ScaleWidth      =   17415
                  TabIndex        =   42
                  Top             =   180
                  Width           =   17415
                  Begin VB.CommandButton CBTransfererVersChargement 
                     BackColor       =   &H0080FF80&
                     Caption         =   "Transférer -> chargement + accès "
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
                     Left            =   11340
                     Style           =   1  'Graphical
                     TabIndex        =   122
                     Top             =   225
                     Width           =   3915
                  End
                  Begin VB.PictureBox PBChoixPosteAnodisationPrevisionnel 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FF8080&
                     ForeColor       =   &H80000008&
                     Height          =   2190
                     Left            =   10800
                     ScaleHeight     =   2160
                     ScaleWidth      =   2805
                     TabIndex        =   116
                     Top             =   1080
                     Visible         =   0   'False
                     Width           =   2835
                     Begin VB.CommandButton CBChoixPosteAnodisationPrevisionnel 
                        BackColor       =   &H00FFFFC0&
                        Caption         =   "C16 IMPOSE"
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
                        Index           =   4
                        Left            =   60
                        Style           =   1  'Graphical
                        TabIndex        =   121
                        Top             =   1740
                        Visible         =   0   'False
                        Width           =   2715
                     End
                     Begin VB.CommandButton CBChoixPosteAnodisationPrevisionnel 
                        BackColor       =   &H00FFFFC0&
                        Caption         =   "C15 IMPOSE"
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
                        Left            =   60
                        Style           =   1  'Graphical
                        TabIndex        =   120
                        Top             =   1320
                        Width           =   2715
                     End
                     Begin VB.CommandButton CBChoixPosteAnodisationPrevisionnel 
                        BackColor       =   &H00FFFFC0&
                        Caption         =   "C14 IMPOSE"
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
                        Left            =   60
                        Style           =   1  'Graphical
                        TabIndex        =   119
                        Top             =   900
                        Width           =   2715
                     End
                     Begin VB.CommandButton CBChoixPosteAnodisationPrevisionnel 
                        BackColor       =   &H00FFFFC0&
                        Caption         =   "C13 IMPOSE"
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
                        Left            =   60
                        Style           =   1  'Graphical
                        TabIndex        =   118
                        Top             =   480
                        Width           =   2715
                     End
                     Begin VB.CommandButton CBChoixPosteAnodisationPrevisionnel 
                        BackColor       =   &H00FFFFC0&
                        Caption         =   "Automatique"
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
                        Index           =   0
                        Left            =   60
                        Style           =   1  'Graphical
                        TabIndex        =   117
                        Top             =   60
                        Width           =   2715
                     End
                  End
                  Begin VB.CommandButton CBTransfererVersChargement 
                     BackColor       =   &H0080FF80&
                     Caption         =   "Transférer -> chargement"
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
                     Left            =   7260
                     Style           =   1  'Graphical
                     TabIndex        =   115
                     Top             =   225
                     Width           =   3915
                  End
                  Begin MSComctlLib.Toolbar TOBGestionGrillePrevisionnel 
                     Height          =   405
                     Index           =   0
                     Left            =   720
                     TabIndex        =   52
                     Top             =   240
                     Width           =   6375
                     _ExtentX        =   11245
                     _ExtentY        =   714
                     ButtonWidth     =   2514
                     ButtonHeight    =   661
                     AllowCustomize  =   0   'False
                     Wrappable       =   0   'False
                     Style           =   1
                     ImageList       =   "ILOutilsGestionGrilles2"
                     HotImageList    =   "ILOutilsGestionGrilles2"
                     _Version        =   393216
                     BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                        NumButtons      =   6
                        BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Key             =   "SupprimerLigne"
                           Object.ToolTipText     =   " Supprime une ligne sur une grille "
                           ImageIndex      =   1
                        EndProperty
                        BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Style           =   3
                        EndProperty
                        BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Key             =   "CompacterGrille"
                           Object.ToolTipText     =   " Compacte les lignes d'une grille "
                           ImageIndex      =   2
                        EndProperty
                        BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Style           =   3
                        EndProperty
                        BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Key             =   "InsererLigne"
                           Object.ToolTipText     =   " Insère une ligne dans une grille "
                           ImageIndex      =   3
                        EndProperty
                        BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Style           =   3
                        EndProperty
                     EndProperty
                     BorderStyle     =   1
                  End
                  Begin MSComctlLib.Toolbar TOBGestionGrillePrevisionnel 
                     Height          =   405
                     Index           =   1
                     Left            =   240
                     TabIndex        =   53
                     Top             =   240
                     Width           =   405
                     _ExtentX        =   714
                     _ExtentY        =   714
                     ButtonWidth     =   688
                     ButtonHeight    =   661
                     AllowCustomize  =   0   'False
                     Wrappable       =   0   'False
                     Style           =   1
                     ImageList       =   "ILOutilsGestionGrilles1"
                     HotImageList    =   "ILOutilsGestionGrilles1"
                     _Version        =   393216
                     BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                        NumButtons      =   1
                        BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                           Key             =   "EffacerGrille"
                           ImageIndex      =   1
                        EndProperty
                     EndProperty
                     BorderStyle     =   1
                  End
                  Begin MSMask.MaskEdBox MEBEditionPrevisionnel 
                     Height          =   255
                     Left            =   540
                     TabIndex        =   104
                     Top             =   1260
                     Visible         =   0   'False
                     Width           =   1515
                     _ExtentX        =   2672
                     _ExtentY        =   450
                     _Version        =   393216
                     BorderStyle     =   0
                     Appearance      =   0
                     BackColor       =   16777215
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PromptChar      =   "_"
                  End
                  Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGPrevisionnel 
                     Height          =   10215
                     Left            =   240
                     TabIndex        =   54
                     Top             =   840
                     Width           =   16935
                     _ExtentX        =   29871
                     _ExtentY        =   18018
                     _Version        =   393216
                     BackColor       =   16777215
                     ForeColor       =   0
                     Rows            =   100
                     Cols            =   6
                     BackColorFixed  =   8388608
                     ForeColorFixed  =   16777215
                     BackColorSel    =   16777215
                     BackColorBkg    =   12648447
                     GridColor       =   0
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
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
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
                  Begin VB.Shape SFocusTablePrevisionnel 
                     BorderColor     =   &H000000FF&
                     BorderWidth     =   4
                     Height          =   10230
                     Left            =   240
                     Top             =   840
                     Visible         =   0   'False
                     Width           =   16950
                  End
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGGrilleRecherche 
                  Bindings        =   "FChargementPrevisionnel.frx":21FC
                  Height          =   9135
                  Index           =   1
                  Left            =   120
                  TabIndex        =   114
                  Top             =   1860
                  Width           =   10155
                  _ExtentX        =   17912
                  _ExtentY        =   16113
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "NumGamme"
                  Columns(0).DataField=   "NumGamme"
                  Columns(0).DataWidth=   6
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "RefGamme"
                  Columns(1).DataField=   "RefGamme"
                  Columns(1).DataWidth=   30
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "NomGamme"
                  Columns(2).DataField=   "NomGamme"
                  Columns(2).DataWidth=   50
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   3
                  Splits(0)._UserFlags=   0
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   -1  'True
                  Splits(0)._GSX_SAVERECORDSELECTORS=   0
                  Splits(0).DividerColor=   13160660
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=3"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2566"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2434"
                  Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
                  Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
                  Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(9)=   "Column(2).Width=4366"
                  Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
                  Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   0
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  DefColWidth     =   0
                  HeadLines       =   1
                  FootLines       =   1
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  InsertMode      =   0   'False
                  MultiSelect     =   2
                  DeadAreaBackColor=   13160660
                  RowDividerColor =   13160660
                  RowSubDividerColor=   13160660
                  DirectionAfterEnter=   1
                  DirectionAfterTab=   1
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=750,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=2"
                  _StyleDefs(5)   =   ":id=0,.fontname=Marlett"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
                  _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=975,.italic=0"
                  _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=975,.italic=0"
                  _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
                  _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                  _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                  _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
                  _StyleDefs(48)  =   "Named:id=33:Normal"
                  _StyleDefs(49)  =   ":id=33,.parent=0"
                  _StyleDefs(50)  =   "Named:id=34:Heading"
                  _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(52)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(53)  =   "Named:id=35:Footing"
                  _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(55)  =   "Named:id=36:Selected"
                  _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(57)  =   "Named:id=37:Caption"
                  _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(59)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(61)  =   "Named:id=39:EvenRow"
                  _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(63)  =   "Named:id=40:OddRow"
                  _StyleDefs(64)  =   ":id=40,.parent=33"
                  _StyleDefs(65)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(66)  =   ":id=41,.parent=34"
                  _StyleDefs(67)  =   "Named:id=42:FilterBar"
                  _StyleDefs(68)  =   ":id=42,.parent=33"
               End
               Begin MSAdodcLib.Adodc ADODCGammesAnodisation 
                  Height          =   375
                  Index           =   1
                  Left            =   180
                  Top             =   11160
                  Width           =   10155
                  _ExtentX        =   17912
                  _ExtentY        =   661
                  ConnectMode     =   0
                  CursorLocation  =   3
                  IsolationLevel  =   -1
                  ConnectionTimeout=   15
                  CommandTimeout  =   30
                  CursorType      =   3
                  LockType        =   3
                  CommandType     =   1
                  CursorOptions   =   0
                  CacheSize       =   50
                  MaxRecords      =   0
                  BOFAction       =   0
                  EOFAction       =   0
                  ConnectStringType=   1
                  Appearance      =   1
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Orientation     =   0
                  Enabled         =   -1
                  Connect         =   "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
                  OLEDBString     =   "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
                  OLEDBFile       =   ""
                  DataSourceName  =   ""
                  OtherAttributes =   ""
                  UserName        =   ""
                  Password        =   ""
                  RecordSource    =   $"FChargementPrevisionnel.frx":2224
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _Version        =   393216
               End
               Begin VB.Shape SDecoration 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   2
                  FillColor       =   &H00808080&
                  FillStyle       =   0  'Solid
                  Height          =   1395
                  Index           =   13
                  Left            =   7860
                  Top             =   240
                  Width           =   2415
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   11715
               Index           =   8
               Left            =   29760
               ScaleHeight     =   11655
               ScaleWidth      =   28065
               TabIndex        =   15
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   11715
               Index           =   7
               Left            =   29460
               ScaleHeight     =   11655
               ScaleWidth      =   28065
               TabIndex        =   14
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   11715
               Index           =   5
               Left            =   29160
               ScaleHeight     =   11655
               ScaleWidth      =   28065
               TabIndex        =   13
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   11715
               Index           =   0
               Left            =   45
               ScaleHeight     =   11655
               ScaleWidth      =   28065
               TabIndex        =   12
               Top             =   495
               Width           =   28125
               Begin VB.Frame FNumBarres 
                  Caption         =   " Numéro de barre "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1095
                  Left            =   21480
                  TabIndex        =   135
                  Top             =   7080
                  Visible         =   0   'False
                  Width           =   6375
                  Begin VB.ComboBox ComboBarre 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   360
                     Left            =   840
                     Style           =   2  'Dropdown List
                     TabIndex        =   136
                     Top             =   360
                     Width           =   3135
                  End
               End
               Begin VB.Frame RepositonnerCadre 
                  Caption         =   "Positionner la charge sur un poste:"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   855
                  Left            =   21480
                  TabIndex        =   125
                  Top             =   8160
                  Visible         =   0   'False
                  Width           =   6375
                  Begin VB.ComboBox CBNumPosteDepart 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   360
                     Left            =   840
                     Style           =   2  'Dropdown List
                     TabIndex        =   127
                     Top             =   360
                     Width           =   3135
                  End
                  Begin VB.CommandButton PositionneCharge 
                     BackColor       =   &H8000000E&
                     Caption         =   "Valider"
                     DownPicture     =   "FChargementPrevisionnel.frx":226D
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
                     Left            =   4440
                     MaskColor       =   &H000000FF&
                     Picture         =   "FChargementPrevisionnel.frx":296F
                     TabIndex        =   126
                     Top             =   360
                     UseMaskColor    =   -1  'True
                     Width           =   1575
                  End
               End
               Begin VB.CommandButton CBTransfererVersPrevisionnel 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Transférer -> prévisionnel"
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
                  Left            =   17460
                  Style           =   1  'Graphical
                  TabIndex        =   123
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   3915
               End
               Begin MSMask.MaskEdBox MEBEditionDetailsCharges 
                  Height          =   255
                  Left            =   10785
                  TabIndex        =   38
                  Top             =   900
                  Visible         =   0   'False
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   450
                  _Version        =   393216
                  BorderStyle     =   0
                  Appearance      =   0
                  BackColor       =   16777215
                  ForeColor       =   0
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  PromptChar      =   "_"
               End
               Begin VB.PictureBox PBCriteresRecherche 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00808080&
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
                  Height          =   1515
                  Index           =   0
                  Left            =   180
                  ScaleHeight     =   1485
                  ScaleWidth      =   10125
                  TabIndex        =   94
                  Top             =   180
                  Width           =   10155
                  Begin VB.ComboBox CBRechercherPar 
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
                     Height          =   360
                     Index           =   0
                     ItemData        =   "FChargementPrevisionnel.frx":2C59
                     Left            =   2100
                     List            =   "FChargementPrevisionnel.frx":2C66
                     Style           =   2  'Dropdown List
                     TabIndex        =   102
                     Top             =   120
                     Width           =   3135
                  End
                  Begin VB.TextBox TBContenant 
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
                     Height          =   315
                     Index           =   0
                     Left            =   2100
                     TabIndex        =   98
                     Top             =   1020
                     Width           =   3135
                  End
                  Begin VB.TextBox TBCommencantPar 
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
                     Height          =   315
                     Index           =   0
                     Left            =   2100
                     TabIndex        =   97
                     Top             =   600
                     Width           =   3135
                  End
                  Begin VB.CommandButton CBLancerRecherche 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "RECHERCHER"
                     BeginProperty Font 
                        Name            =   "Small Fonts"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   555
                     Index           =   0
                     Left            =   5460
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":2C9F
                     Style           =   1  'Graphical
                     TabIndex        =   96
                     ToolTipText     =   " Lancer une recherche "
                     Top             =   120
                     UseMaskColor    =   -1  'True
                     Width           =   1755
                  End
                  Begin VB.CommandButton CBRaz 
                     BackColor       =   &H00E0E0E0&
                     Caption         =   "RAZ"
                     BeginProperty Font 
                        Name            =   "Small Fonts"
                        Size            =   6.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   555
                     Index           =   0
                     Left            =   5460
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":2FE1
                     Style           =   1  'Graphical
                     TabIndex        =   95
                     ToolTipText     =   " Annule tris et recherches "
                     Top             =   780
                     UseMaskColor    =   -1  'True
                     Width           =   1755
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Contenant"
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
                     Height          =   255
                     Index           =   26
                     Left            =   840
                     TabIndex        =   101
                     Top             =   1080
                     Width           =   1095
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Commençant par"
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
                     Height          =   255
                     Index           =   25
                     Left            =   180
                     TabIndex        =   100
                     Top             =   660
                     Width           =   1755
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Rechercher par"
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
                     Height          =   255
                     Index           =   24
                     Left            =   0
                     TabIndex        =   99
                     Top             =   180
                     Width           =   1935
                  End
               End
               Begin VB.Frame FRedresseurs 
                  Caption         =   " Redresseurs "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   4215
                  Left            =   10560
                  TabIndex        =   55
                  Top             =   7320
                  Visible         =   0   'False
                  Width           =   10875
                  Begin VB.PictureBox PBPhasesRedresseurs 
                     BackColor       =   &H00C0E0FF&
                     Height          =   3735
                     Left            =   4620
                     ScaleHeight     =   3675
                     ScaleWidth      =   6015
                     TabIndex        =   74
                     Top             =   300
                     Width           =   6075
                     Begin VB.TextBox TBIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   4
                        Left            =   4440
                        MaxLength       =   6
                        TabIndex        =   73
                        Top             =   2460
                        Width           =   855
                     End
                     Begin VB.TextBox TBIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   3
                        Left            =   4440
                        MaxLength       =   6
                        TabIndex        =   70
                        Top             =   1920
                        Width           =   855
                     End
                     Begin VB.TextBox TBIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   2
                        Left            =   4440
                        MaxLength       =   6
                        TabIndex        =   67
                        Top             =   1380
                        Width           =   855
                     End
                     Begin VB.TextBox TBIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   1
                        Left            =   4440
                        MaxLength       =   6
                        TabIndex        =   64
                        Top             =   840
                        Width           =   855
                     End
                     Begin VB.TextBox TBTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   4
                        Left            =   2880
                        MaxLength       =   6
                        TabIndex        =   72
                        Top             =   2460
                        Width           =   855
                     End
                     Begin VB.TextBox TBTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   3
                        Left            =   2880
                        MaxLength       =   6
                        TabIndex        =   69
                        Top             =   1920
                        Width           =   855
                     End
                     Begin VB.TextBox TBTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   2
                        Left            =   2880
                        MaxLength       =   6
                        TabIndex        =   66
                        Top             =   1380
                        Width           =   855
                     End
                     Begin VB.TextBox TBTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Height          =   315
                        Index           =   1
                        Left            =   2880
                        MaxLength       =   6
                        TabIndex        =   63
                        Top             =   840
                        Width           =   855
                     End
                     Begin MSMask.MaskEdBox MEBTempsPhases 
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   62
                        Top             =   840
                        Width           =   855
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _Version        =   393216
                        ClipMode        =   1
                        BackColor       =   16777215
                        MaxLength       =   7
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "#:##:##"
                        PromptChar      =   "_"
                     End
                     Begin MSMask.MaskEdBox MEBTempsPhases 
                        Height          =   315
                        Index           =   2
                        Left            =   1560
                        TabIndex        =   65
                        Top             =   1380
                        Width           =   855
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _Version        =   393216
                        ClipMode        =   1
                        BackColor       =   16777215
                        MaxLength       =   7
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "#:##:##"
                        PromptChar      =   "_"
                     End
                     Begin MSMask.MaskEdBox MEBTempsPhases 
                        Height          =   315
                        Index           =   3
                        Left            =   1560
                        TabIndex        =   68
                        Top             =   1920
                        Width           =   855
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _Version        =   393216
                        ClipMode        =   1
                        BackColor       =   16777215
                        MaxLength       =   7
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "#:##:##"
                        PromptChar      =   "_"
                     End
                     Begin MSMask.MaskEdBox MEBTempsPhases 
                        Height          =   315
                        Index           =   4
                        Left            =   1560
                        TabIndex        =   71
                        Top             =   2460
                        Width           =   855
                        _ExtentX        =   1508
                        _ExtentY        =   556
                        _Version        =   393216
                        ClipMode        =   1
                        BackColor       =   16777215
                        MaxLength       =   7
                        BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                           Name            =   "MS Sans Serif"
                           Size            =   9.75
                           Charset         =   0
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Mask            =   "#:##:##"
                        PromptChar      =   "_"
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H000040C0&
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Total"
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
                        Index           =   23
                        Left            =   480
                        TabIndex        =   91
                        Top             =   3000
                        Width           =   630
                     End
                     Begin VB.Label LTempsTotalGammeRedresseur 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H00C0FFFF&
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
                        Height          =   285
                        Left            =   1560
                        TabIndex        =   90
                        Top             =   3015
                        Width           =   855
                     End
                     Begin VB.Line LDecoration 
                        Index           =   2
                        X1              =   4200
                        X2              =   4200
                        Y1              =   720
                        Y2              =   2880
                     End
                     Begin VB.Line LDecoration 
                        Index           =   1
                        X1              =   1320
                        X2              =   1320
                        Y1              =   660
                        Y2              =   3420
                     End
                     Begin VB.Line LDecoration 
                        Index           =   0
                        X1              =   2640
                        X2              =   2640
                        Y1              =   720
                        Y2              =   2880
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "TEMPS"
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
                        Index           =   12
                        Left            =   1440
                        TabIndex        =   80
                        Top             =   360
                        Width           =   1095
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   495
                        Index           =   9
                        Left            =   1320
                        Shape           =   4  'Rounded Rectangle
                        Top             =   240
                        Width           =   1335
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H000040C0&
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Ph.4"
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
                        Index           =   22
                        Left            =   480
                        TabIndex        =   89
                        Top             =   2460
                        Width           =   630
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H000040C0&
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Ph.3"
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
                        Index           =   21
                        Left            =   480
                        TabIndex        =   88
                        Top             =   1920
                        Width           =   630
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H000040C0&
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Ph.2"
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
                        Index           =   20
                        Left            =   480
                        TabIndex        =   87
                        Top             =   1380
                        Width           =   630
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H000040C0&
                        BorderStyle     =   1  'Fixed Single
                        Caption         =   "Ph.1"
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
                        Index           =   19
                        Left            =   480
                        TabIndex        =   86
                        Top             =   840
                        Width           =   630
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "A"
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
                        Index           =   18
                        Left            =   5400
                        TabIndex        =   85
                        Top             =   2490
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "A"
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
                        Index           =   17
                        Left            =   5400
                        TabIndex        =   84
                        Top             =   1950
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "A"
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
                        Index           =   16
                        Left            =   5400
                        TabIndex        =   83
                        Top             =   1410
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "INTENSITE"
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
                        Index           =   15
                        Left            =   4320
                        TabIndex        =   82
                        Top             =   360
                        Width           =   1335
                     End
                     Begin VB.Label LLibelles 
                        Alignment       =   2  'Center
                        Appearance      =   0  'Flat
                        BackColor       =   &H80000005&
                        BackStyle       =   0  'Transparent
                        Caption         =   "TENSION"
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
                        Index           =   14
                        Left            =   2760
                        TabIndex        =   81
                        Top             =   360
                        Width           =   1335
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "A"
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
                        Index           =   10
                        Left            =   5400
                        TabIndex        =   79
                        Top             =   870
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "V"
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
                        Index           =   9
                        Left            =   3840
                        TabIndex        =   78
                        Top             =   2490
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "V"
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
                        Index           =   8
                        Left            =   3840
                        TabIndex        =   77
                        Top             =   1950
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "V"
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
                        Index           =   7
                        Left            =   3840
                        TabIndex        =   76
                        Top             =   1410
                        Width           =   195
                     End
                     Begin VB.Label LLibelles 
                        BackStyle       =   0  'Transparent
                        Caption         =   "V"
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
                        Index           =   6
                        Left            =   3840
                        TabIndex        =   75
                        Top             =   870
                        Width           =   195
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   555
                        Index           =   2
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   720
                        Width           =   5535
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   555
                        Index           =   4
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1260
                        Width           =   5535
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   555
                        Index           =   5
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1800
                        Width           =   5535
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   555
                        Index           =   7
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   2340
                        Width           =   5535
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   495
                        Index           =   10
                        Left            =   2640
                        Shape           =   4  'Rounded Rectangle
                        Top             =   240
                        Width           =   1575
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   495
                        Index           =   11
                        Left            =   4200
                        Shape           =   4  'Rounded Rectangle
                        Top             =   240
                        Width           =   1575
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   555
                        Index           =   12
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   2880
                        Width           =   2415
                     End
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00004000&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "En U ou I"
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
                     Index           =   2
                     Left            =   3150
                     TabIndex        =   59
                     Top             =   300
                     Width           =   1170
                  End
                  Begin VB.Image IPhasesAnodisation 
                     Height          =   2010
                     Left            =   240
                     Picture         =   "FChargementPrevisionnel.frx":31D3
                     Top             =   600
                     Width           =   2925
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
                     Left            =   240
                     TabIndex        =   58
                     Top             =   300
                     Width           =   2910
                  End
                  Begin VB.Label LModeUouI 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "En U"
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
                     Index           =   0
                     Left            =   3270
                     TabIndex        =   57
                     Top             =   720
                     Width           =   915
                  End
                  Begin VB.Label LModeUouI 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "En I"
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
                     Index           =   1
                     Left            =   3270
                     TabIndex        =   56
                     Top             =   1140
                     Width           =   915
                  End
                  Begin VB.Shape SDecoration 
                     BorderWidth     =   2
                     FillColor       =   &H00FFFFC0&
                     FillStyle       =   0  'Solid
                     Height          =   960
                     Index           =   1
                     Left            =   3150
                     Top             =   615
                     Width           =   1170
                  End
               End
               Begin VB.PictureBox PBReferencesClient 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FF8080&
                  ForeColor       =   &H80000008&
                  Height          =   3885
                  Left            =   15000
                  ScaleHeight     =   3855
                  ScaleWidth      =   5775
                  TabIndex        =   33
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   5805
                  Begin VB.CommandButton CBAnnulerReferencesClient 
                     BackColor       =   &H00FFFFC0&
                     Caption         =   "&Annuler"
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
                     Left            =   3420
                     Style           =   1  'Graphical
                     TabIndex        =   35
                     Top             =   3360
                     Width           =   1995
                  End
                  Begin VB.CommandButton CBValiderReferencesClient 
                     BackColor       =   &H00FFFFC0&
                     Caption         =   "&Valider "
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
                     Left            =   360
                     Style           =   1  'Graphical
                     TabIndex        =   34
                     Top             =   3360
                     Width           =   1995
                  End
                  Begin MSMask.MaskEdBox MEBEditionDetailsReferencesClient 
                     Height          =   255
                     Left            =   180
                     TabIndex        =   36
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _Version        =   393216
                     BorderStyle     =   0
                     Appearance      =   0
                     BackColor       =   16777215
                     ForeColor       =   0
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     PromptChar      =   "_"
                  End
                  Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsReferencesClient 
                     Height          =   3135
                     Left            =   60
                     TabIndex        =   37
                     Top             =   60
                     Width           =   5655
                     _ExtentX        =   9975
                     _ExtentY        =   5530
                     _Version        =   393216
                     BackColor       =   12648447
                     ForeColor       =   0
                     Rows            =   100
                     Cols            =   6
                     BackColorFixed  =   33023
                     ForeColorFixed  =   16777215
                     BackColorSel    =   16777215
                     BackColorBkg    =   16777215
                     GridColor       =   0
                     GridColorFixed  =   0
                     GridColorUnpopulated=   -2147483644
                     WordWrap        =   -1  'True
                     AllowBigSelection=   0   'False
                     FocusRect       =   0
                     HighLight       =   0
                     ScrollBars      =   2
                     AllowUserResizing=   3
                     Appearance      =   0
                     BandDisplay     =   1
                     RowSizingMode   =   1
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
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
               End
               Begin VB.Frame FOptions 
                  Caption         =   " OPTIONS DE LA GAMME "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   4455
                  Left            =   21540
                  TabIndex        =   32
                  Top             =   60
                  Visible         =   0   'False
                  Width           =   6375
                  Begin VB.TextBox EtuveTpsPoste 
                     Height          =   305
                     Left            =   3600
                     TabIndex        =   132
                     Text            =   "20"
                     Top             =   2180
                     Width           =   495
                  End
                  Begin VB.CheckBox CBOptionsEtuve 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Passage dans l'étuve"
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
                     Left            =   420
                     TabIndex        =   128
                     Top             =   2230
                     Width           =   2715
                  End
                  Begin VB.CheckBox CBOptionsPostes 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Activer l'air dans le bain de BRILLANTAGE"
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
                     Left            =   480
                     TabIndex        =   103
                     Top             =   3960
                     Width           =   5475
                  End
                  Begin VB.TextBox TBDelaiSupStabilisationCharge 
                     Alignment       =   1  'Right Justify
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
                     Height          =   360
                     Left            =   3720
                     MaxLength       =   2
                     TabIndex        =   47
                     Top             =   3000
                     Width           =   495
                  End
                  Begin VB.CheckBox CBOptionsPonts 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Forcer la DESCENTE en PETITE VITESSE"
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
                     Left            =   420
                     TabIndex        =   46
                     Top             =   1880
                     Width           =   5475
                  End
                  Begin VB.CheckBox CBOptionsPonts 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Forcer la MONTEE en TRES PETITE VITESSE"
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
                     Left            =   420
                     TabIndex        =   45
                     Top             =   540
                     Width           =   5475
                  End
                  Begin VB.CheckBox CBOptionsPonts 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Forcer la MONTEE en PETITE VITESSE"
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
                     Left            =   420
                     TabIndex        =   44
                     Top             =   900
                     Width           =   5475
                  End
                  Begin VB.CheckBox CBOptionsPonts 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Forcer la DESCENTE en TRES PETITE VITESSE"
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
                     Left            =   420
                     TabIndex        =   43
                     Top             =   1520
                     Width           =   5475
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "minutes"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   210
                     Index           =   33
                     Left            =   4140
                     TabIndex        =   131
                     Top             =   2225
                     Width           =   975
                  End
                  Begin VB.Shape SDecorationActiverAirBainBrillantage 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  'Opaque
                     Height          =   615
                     Left            =   180
                     Shape           =   4  'Rounded Rectangle
                     Top             =   3720
                     Width           =   6015
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "Délai supplémentaire de stabilisation de la charge en ARRET au POSTE"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   735
                     Index           =   13
                     Left            =   480
                     TabIndex        =   49
                     Top             =   2760
                     Width           =   3135
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     BackStyle       =   0  'Transparent
                     Caption         =   "secondes"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Index           =   11
                     Left            =   4440
                     TabIndex        =   48
                     Top             =   3000
                     Width           =   1215
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  'Opaque
                     Height          =   915
                     Index           =   6
                     Left            =   180
                     Shape           =   4  'Rounded Rectangle
                     Top             =   360
                     Width           =   6015
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  'Opaque
                     Height          =   1275
                     Index           =   0
                     Left            =   180
                     Shape           =   4  'Rounded Rectangle
                     Top             =   1300
                     Width           =   6015
                  End
                  Begin VB.Shape SDecoration 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  'Opaque
                     Height          =   975
                     Index           =   8
                     Left            =   180
                     Shape           =   4  'Rounded Rectangle
                     Top             =   2630
                     Width           =   6015
                  End
               End
               Begin VB.Frame FTempsEtPosteAnodisation 
                  Caption         =   " Temps / compensation / poste d'anodisation"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1035
                  Left            =   10500
                  TabIndex        =   24
                  Top             =   6240
                  Visible         =   0   'False
                  Width           =   10875
                  Begin VB.CommandButton CBCompensation 
                     BackColor       =   &H000000FF&
                     Height          =   315
                     Left            =   2040
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":165DD
                     Style           =   1  'Graphical
                     TabIndex        =   28
                     Top             =   435
                     UseMaskColor    =   -1  'True
                     Width           =   375
                  End
                  Begin VB.TextBox TBCompensation 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   4020
                     MaxLength       =   4
                     TabIndex        =   27
                     Top             =   435
                     Width           =   495
                  End
                  Begin VB.CommandButton CBConfirmationTempsAnodisation 
                     BackColor       =   &H00C0FFFF&
                     Height          =   315
                     Left            =   180
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":168C7
                     Style           =   1  'Graphical
                     TabIndex        =   26
                     Top             =   435
                     UseMaskColor    =   -1  'True
                     Width           =   375
                  End
                  Begin VB.ComboBox CBChoixPosteAnodisation 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   360
                     ItemData        =   "FChargementPrevisionnel.frx":16BB1
                     Left            =   8640
                     List            =   "FChargementPrevisionnel.frx":16BC4
                     Style           =   2  'Dropdown List
                     TabIndex        =   25
                     Top             =   420
                     Width           =   2055
                  End
                  Begin MSMask.MaskEdBox MEBTempsReelAnodisation 
                     Height          =   315
                     Left            =   720
                     TabIndex        =   29
                     Top             =   420
                     Width           =   1035
                     _ExtentX        =   1826
                     _ExtentY        =   556
                     _Version        =   393216
                     ClipMode        =   1
                     BackColor       =   16777215
                     Enabled         =   0   'False
                     MaxLength       =   8
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Mask            =   "##:##:##"
                     PromptChar      =   "_"
                  End
                  Begin VB.Label LLibelles 
                     BackStyle       =   0  'Transparent
                     Caption         =   "+"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Index           =   5
                     Left            =   3780
                     TabIndex        =   61
                     Top             =   420
                     Width           =   195
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     BackStyle       =   0  'Transparent
                     Caption         =   "minutes"
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
                     Index           =   4
                     Left            =   4620
                     TabIndex        =   60
                     Top             =   465
                     Width           =   795
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LTempsAnodisationGamme 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00E0E0E0&
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
                     Height          =   300
                     Left            =   2520
                     TabIndex        =   31
                     Top             =   450
                     Width           =   1200
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Choix du poste d'ANODISATION"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Index           =   3
                     Left            =   5580
                     TabIndex        =   30
                     Top             =   480
                     Width           =   2955
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Shape SDecoration 
                     FillColor       =   &H00FFFFC0&
                     FillStyle       =   0  'Solid
                     Height          =   480
                     Index           =   3
                     Left            =   1920
                     Shape           =   4  'Rounded Rectangle
                     Top             =   360
                     Width           =   3540
                  End
               End
               Begin VB.PictureBox PBEnsembleChargement 
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H00C0E0FF&
                  FillColor       =   &H00FF00FF&
                  Height          =   1995
                  Left            =   21540
                  ScaleHeight     =   1935
                  ScaleWidth      =   4395
                  TabIndex        =   21
                  Top             =   9540
                  Width           =   4455
                  Begin VB.PictureBox PBCharges 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0E0FF&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H80000008&
                     Height          =   855
                     Index           =   2
                     Left            =   2760
                     ScaleHeight     =   855
                     ScaleWidth      =   495
                     TabIndex        =   130
                     Top             =   120
                     Width           =   495
                  End
                  Begin VB.PictureBox PBCharges 
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0E0FF&
                     BorderStyle     =   0  'None
                     ForeColor       =   &H00000000&
                     Height          =   855
                     Index           =   1
                     Left            =   3600
                     ScaleHeight     =   855
                     ScaleWidth      =   435
                     TabIndex        =   129
                     Top             =   120
                     Width           =   435
                  End
                  Begin VB.Image IAutorisationChargement 
                     Height          =   1575
                     Left            =   1040
                     Picture         =   "FChargementPrevisionnel.frx":16C05
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   675
                  End
                  Begin VB.Image IEtatsPostes 
                     Appearance      =   0  'Flat
                     BorderStyle     =   1  'Fixed Single
                     Height          =   315
                     Index           =   2
                     Left            =   2640
                     Picture         =   "FChargementPrevisionnel.frx":1A40F
                     Top             =   1440
                     Width           =   735
                  End
                  Begin VB.Image IEtatsPostes 
                     Appearance      =   0  'Flat
                     BorderStyle     =   1  'Fixed Single
                     Height          =   315
                     Index           =   1
                     Left            =   3480
                     Picture         =   "FChargementPrevisionnel.frx":1AF01
                     Top             =   1440
                     Width           =   735
                  End
                  Begin VB.Label LNomsPostes 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Chgt2"
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
                     Left            =   2640
                     TabIndex        =   23
                     Top             =   1140
                     Width           =   735
                  End
                  Begin VB.Label LNomsPostes 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "Chgt1"
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
                     Left            =   3480
                     TabIndex        =   22
                     Top             =   1140
                     Width           =   735
                  End
               End
               Begin VB.Frame FGammeAnodisation 
                  Caption         =   " Gamme "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   795
                  Left            =   10500
                  TabIndex        =   16
                  Top             =   5400
                  Width           =   10875
                  Begin VB.TextBox TBMatiere 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   360
                     Left            =   9000
                     MaxLength       =   6
                     TabIndex        =   133
                     Top             =   275
                     Width           =   1695
                  End
                  Begin VB.TextBox TBNumGammeAnodisation 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   315
                     Left            =   1740
                     MaxLength       =   6
                     TabIndex        =   18
                     Top             =   300
                     Width           =   1215
                  End
                  Begin VB.CommandButton CBRechercheGamme 
                     Height          =   315
                     Left            =   3120
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargementPrevisionnel.frx":1B9F3
                     Style           =   1  'Graphical
                     TabIndex        =   17
                     ToolTipText     =   " Lancer une recherche "
                     Top             =   300
                     UseMaskColor    =   -1  'True
                     Width           =   315
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "matière"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Index           =   31
                     Left            =   8040
                     TabIndex        =   134
                     Top             =   330
                     Width           =   705
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "N° de la gamme"
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
                     Left            =   120
                     TabIndex        =   20
                     Top             =   330
                     Width           =   1515
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LNomGammeAnodisation 
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
                     Left            =   3600
                     TabIndex        =   19
                     Top             =   300
                     Width           =   4335
                  End
               End
               Begin MSComctlLib.Toolbar TOBGestionGrilleChargement 
                  Height          =   405
                  Index           =   0
                  Left            =   10980
                  TabIndex        =   50
                  Top             =   180
                  Width           =   6360
                  _ExtentX        =   11218
                  _ExtentY        =   714
                  ButtonWidth     =   2514
                  ButtonHeight    =   661
                  AllowCustomize  =   0   'False
                  Wrappable       =   0   'False
                  Style           =   1
                  ImageList       =   "ILOutilsGestionGrilles2"
                  HotImageList    =   "ILOutilsGestionGrilles2"
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   6
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "SupprimerLigne"
                        Object.ToolTipText     =   " Supprime une ligne sur une grille "
                        ImageIndex      =   1
                     EndProperty
                     BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Style           =   3
                     EndProperty
                     BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "CompacterGrille"
                        Object.ToolTipText     =   " Compacte les lignes d'une grille "
                        ImageIndex      =   2
                     EndProperty
                     BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Style           =   3
                     EndProperty
                     BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "InsererLigne"
                        Object.ToolTipText     =   " Insère une ligne dans une grille "
                        ImageIndex      =   3
                     EndProperty
                     BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Style           =   3
                     EndProperty
                  EndProperty
                  BorderStyle     =   1
               End
               Begin MSComctlLib.Toolbar TOBGestionGrilleChargement 
                  Height          =   405
                  Index           =   1
                  Left            =   10500
                  TabIndex        =   51
                  Top             =   180
                  Width           =   405
                  _ExtentX        =   714
                  _ExtentY        =   714
                  ButtonWidth     =   688
                  ButtonHeight    =   661
                  AllowCustomize  =   0   'False
                  Wrappable       =   0   'False
                  Style           =   1
                  ImageList       =   "ILOutilsGestionGrilles1"
                  HotImageList    =   "ILOutilsGestionGrilles1"
                  _Version        =   393216
                  BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
                     NumButtons      =   1
                     BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                        Key             =   "EffacerGrille"
                        ImageIndex      =   1
                     EndProperty
                  EndProperty
                  BorderStyle     =   1
               End
               Begin TrueOleDBGrid80.TDBGrid TDBGGrilleRecherche 
                  Bindings        =   "FChargementPrevisionnel.frx":1BD35
                  Height          =   9135
                  Index           =   0
                  Left            =   180
                  TabIndex        =   93
                  Top             =   1860
                  Width           =   10155
                  _ExtentX        =   17912
                  _ExtentY        =   16113
                  _LayoutType     =   4
                  _RowHeight      =   -2147483647
                  _WasPersistedAsPixels=   0
                  Columns(0)._VlistStyle=   0
                  Columns(0)._MaxComboItems=   5
                  Columns(0).Caption=   "NumGamme"
                  Columns(0).DataField=   "NumGamme"
                  Columns(0).DataWidth=   6
                  Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(1)._VlistStyle=   0
                  Columns(1)._MaxComboItems=   5
                  Columns(1).Caption=   "RefGamme"
                  Columns(1).DataField=   "RefGamme"
                  Columns(1).DataWidth=   30
                  Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns(2)._VlistStyle=   0
                  Columns(2)._MaxComboItems=   5
                  Columns(2).Caption=   "NomGamme"
                  Columns(2).DataField=   "NomGamme"
                  Columns(2).DataWidth=   50
                  Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
                  Columns.Count   =   3
                  Splits(0)._UserFlags=   0
                  Splits(0).RecordSelectorWidth=   503
                  Splits(0)._SavedRecordSelectors=   -1  'True
                  Splits(0)._GSX_SAVERECORDSELECTORS=   0
                  Splits(0).DividerColor=   13160660
                  Splits(0).SpringMode=   0   'False
                  Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
                  Splits(0)._ColumnProps(0)=   "Columns.Count=3"
                  Splits(0)._ColumnProps(1)=   "Column(0).Width=2566"
                  Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
                  Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2434"
                  Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
                  Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
                  Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
                  Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
                  Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
                  Splits(0)._ColumnProps(9)=   "Column(2).Width=4366"
                  Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
                  Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
                  Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
                  Splits.Count    =   1
                  PrintInfos(0)._StateFlags=   0
                  PrintInfos(0).Name=   "piInternal 0"
                  PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
                  PrintInfos(0).PageHeaderHeight=   0
                  PrintInfos(0).PageFooterHeight=   0
                  PrintInfos.Count=   1
                  DefColWidth     =   0
                  HeadLines       =   1
                  FootLines       =   1
                  MultipleLines   =   0
                  CellTipsWidth   =   0
                  InsertMode      =   0   'False
                  MultiSelect     =   2
                  DeadAreaBackColor=   13160660
                  RowDividerColor =   13160660
                  RowSubDividerColor=   13160660
                  DirectionAfterEnter=   1
                  DirectionAfterTab=   1
                  MaxRows         =   250000
                  ViewColumnCaptionWidth=   0
                  ViewColumnWidth =   0
                  _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
                  _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
                  _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
                  _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
                  _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=750,.italic=0"
                  _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=2"
                  _StyleDefs(5)   =   ":id=0,.fontname=Marlett"
                  _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
                  _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
                  _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
                  _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=975,.italic=0"
                  _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
                  _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=975,.italic=0"
                  _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
                  _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
                  _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
                  _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
                  _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
                  _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
                  _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
                  _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
                  _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
                  _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
                  _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
                  _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
                  _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
                  _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
                  _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
                  _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
                  _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
                  _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
                  _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
                  _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
                  _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
                  _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
                  _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
                  _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
                  _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
                  _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
                  _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
                  _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
                  _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
                  _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
                  _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
                  _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
                  _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
                  _StyleDefs(48)  =   "Named:id=33:Normal"
                  _StyleDefs(49)  =   ":id=33,.parent=0"
                  _StyleDefs(50)  =   "Named:id=34:Heading"
                  _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(52)  =   ":id=34,.wraptext=-1"
                  _StyleDefs(53)  =   "Named:id=35:Footing"
                  _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
                  _StyleDefs(55)  =   "Named:id=36:Selected"
                  _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(57)  =   "Named:id=37:Caption"
                  _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
                  _StyleDefs(59)  =   "Named:id=38:HighlightRow"
                  _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
                  _StyleDefs(61)  =   "Named:id=39:EvenRow"
                  _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
                  _StyleDefs(63)  =   "Named:id=40:OddRow"
                  _StyleDefs(64)  =   ":id=40,.parent=33"
                  _StyleDefs(65)  =   "Named:id=41:RecordSelector"
                  _StyleDefs(66)  =   ":id=41,.parent=34"
                  _StyleDefs(67)  =   "Named:id=42:FilterBar"
                  _StyleDefs(68)  =   ":id=42,.parent=33"
               End
               Begin MSAdodcLib.Adodc ADODCGammesAnodisation 
                  Height          =   375
                  Index           =   0
                  Left            =   180
                  Top             =   11160
                  Width           =   10155
                  _ExtentX        =   17912
                  _ExtentY        =   661
                  ConnectMode     =   0
                  CursorLocation  =   3
                  IsolationLevel  =   -1
                  ConnectionTimeout=   15
                  CommandTimeout  =   30
                  CursorType      =   3
                  LockType        =   3
                  CommandType     =   1
                  CursorOptions   =   0
                  CacheSize       =   50
                  MaxRecords      =   0
                  BOFAction       =   0
                  EOFAction       =   0
                  ConnectStringType=   1
                  Appearance      =   1
                  BackColor       =   -2147483643
                  ForeColor       =   -2147483640
                  Orientation     =   0
                  Enabled         =   -1
                  Connect         =   "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
                  OLEDBString     =   "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
                  OLEDBFile       =   ""
                  DataSourceName  =   ""
                  OtherAttributes =   ""
                  UserName        =   ""
                  Password        =   ""
                  RecordSource    =   $"FChargementPrevisionnel.frx":1BD5D
                  Caption         =   ""
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  _Version        =   393216
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsCharges 
                  Height          =   4515
                  Left            =   10500
                  TabIndex        =   39
                  Top             =   780
                  Width           =   10875
                  _ExtentX        =   19182
                  _ExtentY        =   7964
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   0
                  Rows            =   100
                  Cols            =   6
                  BackColorFixed  =   128
                  ForeColorFixed  =   16777215
                  BackColorSel    =   16777215
                  BackColorBkg    =   12648447
                  GridColor       =   0
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
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
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
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  BackColor       =   &H000080FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Choix du poste de chargement"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000FFFF&
                  Height          =   375
                  Index           =   1
                  Left            =   21540
                  TabIndex        =   40
                  Top             =   9180
                  Width           =   4455
               End
               Begin VB.Shape SFocusTableDetailsCharges 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   4530
                  Left            =   10500
                  Top             =   780
                  Visible         =   0   'False
                  Width           =   10890
               End
            End
         End
      End
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   28020
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   11910
      Width           =   28080
      Begin MSComctlLib.ImageList ILImagesNumChoix 
         Left            =   6120
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   46
         ImageHeight     =   14
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1BDA6
               Key             =   "choix 1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1C5A0
               Key             =   "choix 2"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1CD9A
               Key             =   "choix 3"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1D594
               Key             =   "choix 4"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1DD8E
               Key             =   "choix 5"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1E588
               Key             =   "choix 6"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1ED82
               Key             =   "choix 7"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1F57C
               Key             =   "choix 8"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1FD76
               Key             =   "choix 9"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":20570
               Key             =   "choix 10"
            EndProperty
         EndProperty
      End
      Begin VB.Timer TimerCalculPrevisionnel 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   7740
         Top             =   120
      End
      Begin MSComctlLib.ImageList ILImagesColorations 
         Left            =   5400
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":20D6A
               Key             =   "anodisation"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":210BC
               Key             =   "spectro"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":2140E
               Key             =   "or"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":21760
               Key             =   "noir"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FChargementPrevisionnel.frx":21AB2
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
         Left            =   24600
         MaskColor       =   &H00FF00FF&
         Picture         =   "FChargementPrevisionnel.frx":221B4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBReduire 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Réduire la fenêtre"
         DownPicture     =   "FChargementPrevisionnel.frx":228B6
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   22320
         MaskColor       =   &H00FF00FF&
         Picture         =   "FChargementPrevisionnel.frx":22FB8
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   " Réduire cette fenêtre à la taille minimum "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2115
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FChargementPrevisionnel.frx":236BA
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
            Picture         =   "FChargementPrevisionnel.frx":23864
            Style           =   1  'Graphical
            TabIndex        =   8
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
            TabIndex        =   7
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   6
            Top             =   720
            Width           =   915
         End
      End
      Begin VB.Timer TimerChargementPrevisionnel 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   7200
         Top             =   120
      End
      Begin MSComctlLib.ImageList ILOutilsDivers 
         Left            =   2040
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   47
         ImageHeight     =   19
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":23A0E
               Key             =   "croix de condamnation"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":24510
               Key             =   "rectangle vert"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILIcones 
         Left            =   4020
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   9
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":25012
               Key             =   "fleche haut"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":25C66
               Key             =   "fleche basse"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":268BA
               Key             =   "fleche gauche"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":2750E
               Key             =   "fleche droite"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":28162
               Key             =   "fleche haut ombre"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":28DB6
               Key             =   "fleche gauche ombre"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":29A0A
               Key             =   "fleche droite ombre"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":2A65E
               Key             =   "sens interdit"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":2B2B2
               Key             =   "etoile"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILImagesPourGrilles 
         Left            =   4680
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   14
         ImageHeight     =   15
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":2BF06
               Key             =   "bouton bas"
            EndProperty
         EndProperty
      End
      Begin PicClip.PictureClip PCCharges 
         Left            =   9300
         Top             =   60
         _ExtentX        =   7673
         _ExtentY        =   17463
         _Version        =   393216
         Rows            =   12
         Cols            =   10
         Picture         =   "FChargementPrevisionnel.frx":2C22E
      End
      Begin MSComctlLib.ImageList ILOutilsGestionGrilles2 
         Left            =   3360
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   88
         ImageHeight     =   19
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":B8AA0
               Key             =   "supprimer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":B9E8A
               Key             =   "compacter"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":BB274
               Key             =   "inserer"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILOutilsGestionGrilles1 
         Left            =   2700
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   19
         ImageHeight     =   19
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":BC65E
               Key             =   "effacer grille"
            EndProperty
         EndProperty
      End
      Begin PicClip.PictureClip PCBarres 
         Left            =   6600
         Top             =   0
         _ExtentX        =   7673
         _ExtentY        =   17463
         _Version        =   393216
         Rows            =   12
         Cols            =   10
         Picture         =   "FChargementPrevisionnel.frx":BCB24
      End
      Begin MSComctlLib.ImageList ILGrillesDonnees 
         Left            =   1440
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
               Picture         =   "FChargementPrevisionnel.frx":149396
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1495A2
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1497AE
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":1499BA
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":149BC6
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":149DD2
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":149FDE
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14A1EA
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14A3F6
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14A602
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14A80E
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14AA1A
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14AC26
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14AE32
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14B03E
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14B24A
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14B456
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14B65A
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14B85E
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14BA62
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14BC66
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14BE6A
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14C06E
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14C272
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14C476
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargementPrevisionnel.frx":14C67A
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin VB.Label LDonneesTransmisesAPI 
         Alignment       =   2  'Center
         BackColor       =   &H000000C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   19260
         TabIndex        =   92
         Top             =   180
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   465
         Left            =   8340
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FChargementPrevisionnel.frx":14C87E
      ScaleHeight     =   315
      ScaleWidth      =   28020
      TabIndex        =   1
      Top             =   0
      Width           =   28080
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGEMENT ET PREVISIONNEL"
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
         Left            =   3780
         TabIndex        =   2
         Top             =   60
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "FChargementPrevisionnel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant le chargement et le prévisionnel
' Nom                    : FChargementPrevisionnel.frm
' Date de création : 27/01/2011
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const NBR_COLONNES_DETAILS_CHARGES  As Integer = 7
Private Const NBR_COLONNES_DETAILS_REFERENCES_CLIENT  As Integer = 2

Private Const NBR_COLONNES_PREVISIONNEL  As Integer = 1

Private Const NBR_LIGNES_MOTEUR_INFERENCE As Integer = 30
Private Const NBR_COLONNES_MOTEUR_INFERENCE As Integer = 3
    
Private Const IMG_BOUTON_BAS As String = "bouton bas"
Private Const IMG_ANODISATION As String = "anodisation"
Private Const IMG_SPECTRO As String = "spectro"
Private Const IMG_OR As String = "or"
Private Const IMG_NOIR As String = "noir"

Private Const TITRE_FENETRE As String = "CHARGEMENT ET PREVISIONNEL"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---

'--- onglets ---
Public Enum ONGLETS_CHARGEMENT_PREVISIONNEL
    O_CHARGEMENT = 0
    O_PREVISIONNEL = 1
End Enum

Private Enum COLONNES_GRILLE_RECHERCHE
    C_NUM_GAMME = 0
    C_REF_GAMME = 1
    C_NOM_GAMME = 2
End Enum

Private Enum IDX_RECHERCHER_PAR
    IDX_NUM_GAMME = 1
    IDX_REF_GAMME = 2
    IDX_NOM_GAMME = 3
End Enum

Private Enum TYPES_GRILLES
    TG_CHARGEMENT = 0
    TG_PREVISIONNEL = 1
End Enum

Private Enum COLONNES_DETAILS_CHARGES
    C_NUM_LIGNES = 0
    C_NUM_COMMANDE_INTERNE = 1           'n° de commande interne
    'C_NBR_REPARATIONS = 2                         'nombre de réparations
    C_CODE_CLIENT = 2                                   'code du client
    C_NBR_PIECES = 3                                    'nombre de pièces
    C_DESIGNATION = 4                                   'désignation
    C_GAMME = 5
    C_MATIERE = 6
    C_OBSERVATIONS = 7                                'observations
                                               'matière
End Enum

Private Enum COLONNES_DETAILS_REFERENCES_CLIENT
    C_NUM_LIGNES = 0
    C_NBR_PIECES = 1                                    'nombre de pièces
    C_REFERENCE_CLIENT = 2                       'référence donnée par le client
End Enum

Private Enum IDX_REDRESSEURS
    IDX_REDRESSEUR_C13_A_C16 = 1
End Enum

Private Enum COLONNES_PREVISIONNEL
    C_NUM_LIGNES = 0
    C_CHOIX_IA = 1                                          'meilleur choix pour l'entrée dans la ligne (moteur d'inférence)
    C_NUM_COMMANDE_INTERNE = 2          'n° de commande interne
    C_NBR_REPARATIONS = 3                        'nombre de réparations
    C_CODE_CLIENT = 4                                  'code du client
    C_NBR_PIECES = 5                                    'nombre de pièces
    C_DESIGNATION = 6                                   'désignation
    C_NUM_BARRE = 7                                    'n° de barre
    C_NUM_GAMME_ANODISATION = 8          'n° de la gamme d'anodisation
    C_PASSAGE_ANODISATION = 9                 'indique un passage dans un des bains d'anodisation
    C_PASSAGE_SPECTRO = 10                      'indique un passage dans le bain de spectrocoloration
    C_PASSAGE_OR = 11                                 'indique un passage dans le bain d'or
    C_PASSAGE_NOIR = 12                             'indique un passage dans le bain de noir
    C_CHOIX_POSTE_ANODISATION = 13       'choix du poste d'anodisation
End Enum

Private Enum COLONNES_MOTEUR_INFERENCE
    C_NUM_LIGNES = 0         'n° de ligne
    C_DONNEES_1 = 1           'données 1
    C_DONNEES_2 = 2           'données 2
    C_DONNEES_3 = 3           'données 3
End Enum

'--- types privées ---
Private Type ImgDetailsReferencesClient
    NbrPieces As Double                                'nombre de pièces
    ReferenceClient As String                      'référence du client
End Type

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                            'pour interdire certains évènements
Private GammeTrouvee As Boolean                                    'TRUE = indique que la gamme a été trouvé
Private GammePassableEnLigne As Boolean                     'TRUE = indique que la gamme est passable en ligne
Private BoutonCalculerALeFocus As Boolean                      'le bouton calculer à le focus

Private MemNumLigne As Integer                                        'mémoire d'un n° de ligne dans une des grilles
Private MemNumColonne As Integer                                    'mémoire d'un n° de colonne dans une des grilles
Private MemNumLigneDetailsCharges As Integer               'mémoire d'un n° de ligne dans la grille des détails des charges
Private MemNumColonneDetailsCharges As Integer           'mémoire d'un n° de colonne dans la grille des détails des charges

Private LigneDepartDeplacement As Integer                       'ligne de départ en cas de déplacement d'un détail

Private MemDernierBouton As Long                                    'mémoire du dernier bouton

Private MemTempsReelAnodisationSecondes As Long      'mémorisation du temps réel d'anodisation en secondes
Private MemTempsReelAnodisationTexte As String            'mémorisation du temps réel d'anodisation en texte

Private NumBarreEnCours As Integer                                   'représente un numéro du montage en cours
Private ModeUouIEnCours As MODES_U_OU_I                   'mode U ou I en cours

'--- tableaux privés ---
'ATTENTION les tableaux TChargement et TPrevisionnel sont communs à tout le programme
'afin d'être disponible par le moteur d'inférence
Private TDetailsReferencesClient(1 To NBR_LIGNES_DETAILS_REFERENCES_CLIENT) As ImgDetailsReferencesClient

'--- variables publiques ---
Public NumFenetre As Long                                                 'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBAnnulerReferencesClient_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- focus ---
    SFocusTableDetailsCharges.Visible = False
    
    '--- rendre le contrôle invisible ---
    PBReferencesClient.Visible = False
    
End Sub


Private Sub EtuveTpsPoste_Change()
  
  If Not IsNumeric(EtuveTpsPoste.Text) Then
    MsgBox ("Rentrer une valeur numérique")
    EtuveTpsPoste.Text = 20
  End If
End Sub

Private Sub PositionneCharge_Click()
  
    
   If CBNumPosteDepart.ListIndex >= 0 Then
    IntroductionChargeAuChargement (CBNumPosteDepart.ListIndex + 1)
   Else
    
    Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_501)
   
   End If
   
  
  

    
End Sub
Private Sub CBCalculerPrevisionnel_Click()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- compression de la grille et affichage ---
    GestionPrevisionnel GG_COMPRESSION
    GestionPrevisionnel GG_AFFICHAGE

    '--- calcul du prévisionnel avec affichage des choix ---
    CalculPrevisionnelAvecAffichageChoix
    
End Sub

Private Sub CBCalculerPrevisionnel_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    BoutonCalculerALeFocus = True

End Sub

Private Sub CBCalculerPrevisionnel_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    BoutonCalculerALeFocus = False

End Sub

Private Sub CBChoixPosteAnodisationPrevisionnel_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumLigne As Integer, _
           NumColonne As Integer
        
    '--- affectation du numéro de ligne de la grille ---
    NumLigne = MSHFGPrevisionnel.Row
        
    If NumLigne >= 1 And NumLigne <= NBR_LIGNES_PREVISIONNEL Then
        
        With TPrevisionnel(NumLigne)
            
            If .NumCommandeInterne > 0 Then
                
                '--- affectation du choix du poste d'anodisation ---
                .ChoixPosteAnodisation = Index
                
            End If
        
            '--- rafraichir la grille ---
            GestionPrevisionnel GG_AFFICHAGE
            
        End With

    End If

    '--- effacement de la boite de dialogues ---
    PBChoixPosteAnodisationPrevisionnel.Visible = False

End Sub

Private Sub CBCompensation_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim TempsNickeCompense As Long          'temps d'anodisation compensé
    
    If IsNumeric(TBCompensation.Text) = True And MEBTempsReelAnodisation.Enabled = True Then
        
        '--- calcul du temps avec compensation ---
        TempsNickeCompense = MemTempsReelAnodisationSecondes + CInt(TBCompensation.Text) * 60
    
        '--- limite du temps compensé ---
        If TempsNickeCompense < 0 Then TempsNickeCompense = 0
    
        '--- transfert dans le champ du temps d'anodisation ---
        MEBTempsReelAnodisation.Text = CTemps2(TempsNickeCompense)
    
    End If

End Sub

Private Sub CBConfirmationTempsAnodisation_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With MEBTempsReelAnodisation
    
        '--- transfert du temps de gamme ---
        .Text = MemTempsReelAnodisationTexte

        '--- valider le champ d'édition ---
        .Enabled = True

    End With

End Sub

Private Sub CBLancerRecherche_Click(Index As Integer)
    On Error Resume Next
    LanceRechercheOuTri Index
End Sub

Private Sub CBOptionsPonts_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- céchochage des croix inutiles ---
    Select Case Index
        
        Case OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV
            '--- forcer la montée d'une charge en très petite vitesse ---
            If CBOptionsPonts(Index).value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV).value = vbUnchecked
            End If
        
        Case OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV
            '--- forcer la montée d'une charge en petite vitesse ---
            If CBOptionsPonts(Index).value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV).value = vbUnchecked
            End If
        
        Case OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV
            '--- forcer la descente d'une charge en très petite vitesse ---
            If CBOptionsPonts(Index).value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV).value = vbUnchecked
            End If
        
        Case OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV
            '--- forcer la descente d'une charge en petite vitesse ---
            If CBOptionsPonts(Index).value = vbChecked Then
                CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV).value = vbUnchecked
            End If
        
        Case Else
    End Select

End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    DechargeFenetre
End Sub

Private Sub CBRaz_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- vidage des champs / lancement de la requête ---
    With TBCommencantPar(Index)
        .Text = ""
        .Refresh
        .SetFocus
    End With
    With TBContenant(Index)
        .Text = ""
        .Refresh
    End With
    LanceRechercheOuTri Index

End Sub

Private Sub CBRechercheGamme_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la fenêtre des gammes d'anodisation ---
    'les paramêtres sont  TravailSurGrille
    '                                  RechercherPar
    '                                  CommencantPar
    '                                  Contenant
    '                                  MethodeRechercheChoisie
    If TBNumGammeAnodisation.Text <> "" Then
        AppelFenetre FENETRES.F_GAMMES_ANODISATION, False, 1, TBNumGammeAnodisation.Text, "", True
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

Private Sub CBRechercherPar_Click(Index As Integer)
    On Error Resume Next
    If PremiereActivation = True Then
        DoEvents
        CBRaz_Click (Index)
    End If
End Sub

Private Sub CBReduire_Click()
    On Error Resume Next
    Me.WindowState = vbMinimized
End Sub

Private Sub CBReduire_GotFocus()
    
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

Private Sub CBReduire_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBTransfererVersChargement_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim FicheVidePrevisionnel As VarPrevisionnel     'fiche vide prévisionnel
    Dim LigneDepartDeplacement As Integer               'ligne de départ du déplacement
    
    '--- affectation de la ligne de départ du déplacement ---
    LigneDepartDeplacement = MSHFGPrevisionnel.Row
    
    If LigneDepartDeplacement > 0 And LigneDepartDeplacement <= NBR_LIGNES_PREVISIONNEL Then
    
        '--- contrôle des valeurs à la ligne pointée ---
        With TPrevisionnel(LigneDepartDeplacement)
                    
            If .NumCommandeInterne > 0 And .CodeClient <> "" And .NbrPieces > 0 And .Designation <> "" And .NumGammeAnodisation <> "" Then
                    
                '--- effacement complet du chargement ---
                EffacementCompletChargement
                                
                '--- transfert des valeurs dans la zone du chargement ---
                TransfertDonneesEntreGrilles TG_PREVISIONNEL, LigneDepartDeplacement
    
                '--- sélection du n° de barre ---
                ComboBarre.ListIndex = TPrevisionnel(LigneDepartDeplacement).NumBarre
                
                Call ComboBarre_Click
                
                '--- affectation du numéro de gamme et chargement de la gamme dans la mémoire ---
                TBNumGammeAnodisation.Text = TPrevisionnel(LigneDepartDeplacement).NumGammeAnodisation
                ChargeGammeAnodisationChargement TBNumGammeAnodisation.Text
                
                '--- affectation du choix du poste d'anodisation ---
                DoEvents
                CBChoixPosteAnodisation.ListIndex = TPrevisionnel(LigneDepartDeplacement).ChoixPosteAnodisation

                '--- effacement de la ligne ---
                TPrevisionnel(LigneDepartDeplacement) = FicheVidePrevisionnel
                
                '--- compression de la grille ---
                GestionPrevisionnel GG_COMPRESSION
                GestionPrevisionnel GG_AFFICHAGE
            
            End If
    
        End With
    
    End If
    
    '--- changement de l'onglet en fonction du bouton cliqué ---
    If Index = 1 Then
        CTOnglets.CurrTab = ONGLETS_CHARGEMENT_PREVISIONNEL.O_CHARGEMENT
    End If

End Sub

Private Sub CBValiderReferencesClient_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- mémorisation ---
    GestionDetailsReferencesClient GG_MEMORISATION
    
    '--- focus ---
    SFocusTableDetailsCharges.Visible = False
    
    '--- rendre le contrôle invisible ---
    PBReferencesClient.Visible = False
    
    '--- réaffichage ---
    GestionDetailsCharges GG_AFFICHAGE
    
End Sub

Private Sub CTOnglets_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case CTOnglets.CurrTab

        Case ONGLETS_CHARGEMENT_PREVISIONNEL.O_CHARGEMENT
            '--- chargement ---
            With MSHFGDetailsCharges
                If .Visible = True Then
                    .SetFocus
                End If
            End With
        
        Case ONGLETS_CHARGEMENT_PREVISIONNEL.O_PREVISIONNEL
            '--- prévisionnel ---
            With CBCalculerPrevisionnel
                If .Visible = True Then
                    .SetFocus
                End If
            End With

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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    CBQuitter_Click
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
    End If
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Left = -HSDeplacementFenetre.value
End Sub

Private Sub IAutorisationChargement_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case Source.Name
        
        Case IAutorisationChargement.Name
            '--- agissement à l'intèrieur du contrôle de départ ---
            Select Case State
     
                Case 0
                    '--- 0 = Entre (le contrôle source entre dans la portée de la cible) ---
                    Source.DragIcon = ILIcones.ListImages("fleche droite ombre").ExtractIcon
     
                Case 1
                    '--- 1 = Sort (le contrôle source sort de la portée de la cible) ---
                    Source.DragIcon = ILIcones.ListImages("sens interdit").ExtractIcon
                
                Case 2
                    '--- 2 = Dessus (le contrôle source est passé d'une ² à une autre dans la cible) ---
                
                Case Else

            End Select
                
        Case Else
                
    End Select

End Sub

Private Sub IAutorisationChargement_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
           
    '--- contrôle du début du glisser/déplacer ---
    If Button = vbKeyLButton Then
        IAutorisationChargement.Drag vbBeginDrag
    End If
    
End Sub

Private Sub LModeUouI_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    If Index = MODES_U_OU_I.M_TENSION Then
        
        '--- changement de couleurs pour la gamme en tension ---
        With LModeUouI(MODES_U_OU_I.M_TENSION)
            .BackColor = COULEURS.ROUGE_3
            .ForeColor = COULEURS.JAUNE_3
        End With
        
        '--- changement de couleurs pour la gamme en courant ---
        With LModeUouI(MODES_U_OU_I.M_INTENSITE)
            .BackColor = COULEURS.BLANC
            .ForeColor = COULEURS.NOIR
        End With
    
        '--- affectation du mode tension ---
        ModeUouIEnCours = MODES_U_OU_I.M_TENSION
    
    Else
        
        '--- changement de couleurs pour la gamme en courant ---
        With LModeUouI(MODES_U_OU_I.M_INTENSITE)
            .BackColor = COULEURS.ROUGE_3
            .ForeColor = COULEURS.JAUNE_3
        End With
        
        '--- changement de couleurs pour la gamme en tension ---
        With LModeUouI(MODES_U_OU_I.M_TENSION)
            .BackColor = COULEURS.BLANC
            .ForeColor = COULEURS.NOIR
        End With
        
        '--- affectation du mode intensité ---
        ModeUouIEnCours = MODES_U_OU_I.M_INTENSITE
        
    End If

End Sub

Private Sub LNomsPostes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    If Button = vbLeftButton Then
        
        '**************************************************************************************************************
        '                                           indication d'une charge prioritaire (clic gauche)
        '**************************************************************************************************************
        With TEtatsPostes(Index)
            If .Condamnation = False Then
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    TEtatsCharges(.NumCharge).ChargePrioritaire = Not (TEtatsCharges(.NumCharge).ChargePrioritaire)
                End If
            End If
        End With
    
    Else
        
        '**************************************************************************************************************
        '                                                   annulation de la charge (clic droit)
        '**************************************************************************************************************
        With TEtatsPostes(Index)
        
            If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                
                '--- demande de confirmation ---
                If AppelFenetre(F_MESSAGE, _
                                        TITRE_MESSAGES, _
                                        MESSAGE_302, _
                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                        TYPES_BOUTONS.T_OUI_NON, _
                                        EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
            
                    '--- effacement de la charge dans l'automate ---
                    If EnvoiNumeroChargePoste(Index, 0) <> OK Then
                        Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                    End If
                
                End If
            
            End If
            
        End With
    
    End If

End Sub

Private Sub ComboBarre_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    
    Dim Index As Integer
    Index = ComboBarre.ListIndex
    
    MsgBox ("click sur barre" & NumBarreEnCours)
    
    NumBarreEnCours = Index
    
End Sub

Private Sub LRenseignementsFenetre_DblClick()
    On Error Resume Next
    If Me.WindowState = vbMaximized Then
        Me.WindowState = vbNormal
    Else
        Me.WindowState = vbMaximized
    End If
End Sub

Private Sub LTempsTotalGammeRedresseur_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer            'pour les boucles FOR...NEXT

    '--- affectation automatique du temps dans la gamme ---
    If InterdireEvenements = False Then
    
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
    
            With TChargement.TGammesAnodisation.TDetailsGammesAnodisation(a)
                        
                '--- recherche de la zone d'anodisation ---
                If Trim(TZones(.NumZone).Codezone) = "C13 à C16" Then
                    
                    '--- affectation dans le tableau ---
                    .TempsAuPosteTexte = "0" & LTempsTotalGammeRedresseur.Caption
                    .TempsAuPosteSecondes = CTempsTexteEnSecondes(.TempsAuPosteTexte)
                    
                    '--- affectation dans les champs ---
                    MEBTempsReelAnodisation.Text = .TempsAuPosteTexte
                    LTempsAnodisationGamme.Caption = .TempsAuPosteTexte
                            
                    '--- mémorisation des temps de l'anodisation ---
                    MemTempsReelAnodisationSecondes = .TempsAuPosteSecondes
                    MemTempsReelAnodisationTexte = .TempsAuPosteTexte
                    
                    Exit For
                
                End If
                            
            End With
    
        Next a

    End If

End Sub

Private Sub MEBEditionDetailsCharges_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    If InterdireEvenements = False Then
    
        '--- rendre invisible la liste des références du client ---
        PBReferencesClient.Visible = False
    
        With MSHFGDetailsCharges
    
            Select Case .Col
        
                Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                    '--- n° de commande interne ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                'Case COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                    '--- nombre de réparations ---
                 '   MemNumLigne = .Row
                  '  MemNumColonne = .Col
                    
                Case COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                    '--- nombre de pièces ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                    
                Case Else
        
            End Select
    
        End With

    End If

End Sub

Private Sub MEBEditionDetailsCharges_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    MemNumLigne = 0
    MemNumColonne = 0
    
    '--- rendre visible le focus de la table ---
    SFocusTableDetailsCharges.Visible = True

End Sub

Private Sub MEBEditionDetailsCharges_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With MSHFGDetailsCharges

        '--- analyse de la touche ---
        Select Case KeyCode

            Case vbKeyDown
                '--- flèche basse ---
                .SetFocus
                If .Row < .Rows - 1 Then .Row = .Row + 1
                KeyCode = 0
            
            Case vbKeyUp
                '--- flèche haute ---
                .SetFocus
                If .Row > .FixedRows Then .Row = .Row - 1
                KeyCode = 0
  
            Case Else
  
        End Select
  
    End With
    
End Sub

Private Sub MEBEditionDetailsCharges_KeyPress(KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    With MSHFGDetailsCharges

        '--- analyse de la touche ---
        Select Case KeyAscii

            Case vbKeyReturn
                '--- touche entrée ---
                Select Case .Col
            
                    Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                        '--- n° de la commande interne ---
                        .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                    
                    'Case COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                        '--- nombre de réparations ---
                    '    .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                
                    Case COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                        '--- nombre de pièces ---
                        If .Row < .Rows - 1 Then .Row = .Row + 1
                        .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                
                    Case Else

                End Select

                '--- mettre le focus sur le tableau ---
                .SetFocus
                KeyAscii = 0

            Case Else
                Select Case .Col
                    Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES, 8
                    'Case COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 1
                    Case COLONNES_DETAILS_CHARGES.C_NBR_PIECES: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6
                    Case Else
                End Select

        End Select

    End With

End Sub

Private Sub MEBEditionDetailsCharges_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim TexteComplet As String, _
            TexteSansMasque As String
    
    '--- analyse des données frappées ---
    If MemNumLigne > 0 And MemNumColonne > 0 Then
    
        '--- affectation ---
        With MEBEditionDetailsCharges
            TexteComplet = .Text
            TexteSansMasque = .ClipText
        End With
      
        Select Case MemNumColonne
        
            Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                '--- n° de commande interne ---
                If TexteSansMasque = "" Then
                    
                    '--- pas de n° de commande interne ---
                    With TChargement.TDetailsCharges(MemNumLigne)
                        .NumCommandeInterne = 0
                        .NumLignesReferencesClient = ""
                    End With
                
                Else
                
                    '--- possibilité de valider plusieurs fois le même numéro de commande interne ---
                    If InsertionCommandeInterne(TG_CHARGEMENT, MemNumLigne, TexteSansMasque) <> TROUVE Then
                        
                        '--- message d'erreur ---
                        MessageErreur TITRE_MESSAGES, MESSAGE_122
                        
                        '--- affectation ---
                        With TChargement.TDetailsCharges(MemNumLigne)
                            .NumCommandeInterne = 0
                            .NumLignesReferencesClient = ""
                        End With
                        
                        '--- replacer le focus sur la grille au bon endroit ---
                        With MSHFGDetailsCharges
                            .Row = MemNumLigne
                            .Col = MemNumColonne
                            .SetFocus
                        End With
                    
                    End If
                
                End If
            
            'Case COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                '--- nombre de réparations ---
             '   If IsNumeric(TexteSansMasque) = True Then
             '       TChargement.TDetailsCharges(MemNumLigne).NbrReparations = TexteSansMasque
             '   Else
             '       TChargement.TDetailsCharges(MemNumLigne).NbrReparations = ""
             '   End If
            
            Case COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                '--- nombre de pièces ---
                If IsNumeric(TexteSansMasque) = True Then
                    TChargement.TDetailsCharges(MemNumLigne).NbrPieces = CDbl(TexteSansMasque)
                Else
                    TChargement.TDetailsCharges(MemNumLigne).NbrPieces = 0
                End If
            
            Case Else
    
        End Select

    End If
    
    '--- focus ---
    SFocusTableDetailsCharges.Visible = False
    
    '--- rendre le contrôle texte invisible ---
    MEBEditionDetailsCharges.Visible = False

    '--- construction de la grille ---
    GestionDetailsCharges GG_AFFICHAGE

End Sub

Private Sub MEBEditionDetailsReferencesClient_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    If InterdireEvenements = False Then
    
        With MSHFGDetailsReferencesClient
    
            Select Case .Col
        
                Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                    '--- nombre de pièces ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                    
                Case COLONNES_DETAILS_REFERENCES_CLIENT.C_REFERENCE_CLIENT
                    '--- référence du client ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                    
                Case Else
        
            End Select
    
        End With

    End If

End Sub

Private Sub MEBEditionDetailsReferencesClient_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    MemNumLigne = 0
    MemNumColonne = 0
    
End Sub

Private Sub MEBEditionDetailsReferencesClient_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With MSHFGDetailsReferencesClient

        '--- analyse de la touche ---
        Select Case KeyCode

            Case vbKeyDown
                '--- flèche basse ---
                .SetFocus
                If .Row < .Rows - 1 Then .Row = .Row + 1
                KeyCode = 0
            
            Case vbKeyUp
                '--- flèche haute ---
                .SetFocus
                If .Row > .FixedRows Then .Row = .Row - 1
                KeyCode = 0
  
            Case Else
  
        End Select
  
    End With

End Sub

Private Sub MEBEditionDetailsReferencesClient_KeyPress(KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    With MSHFGDetailsReferencesClient

        '--- analyse de la touche ---
        Select Case KeyAscii

            Case vbKeyReturn
                '--- touche entrée ---
                Select Case .Col
            
                    Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                        '--- nombre de pièces ---
                        If .Row < .Rows - 1 Then .Row = .Row + 1
                        .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                    
                    Case Else

                End Select

                '--- mettre le focus sur le tableau ---
                .SetFocus
                KeyAscii = 0

            Case Else
                Select Case .Col
                    Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6
                    Case Else
                End Select

        End Select

    End With

End Sub

Private Sub MEBEditionDetailsReferencesClient_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim TexteComplet As String, _
            TexteSansMasque As String
    
    '--- analyse des données frappées ---
    If MemNumLigne > 0 And MemNumColonne > 0 Then
    
        '--- affectation ---
        With MEBEditionDetailsReferencesClient
            TexteComplet = .Text
            TexteSansMasque = .ClipText
        End With
        
        Select Case MemNumColonne
    
            Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                '--- nombre de pièces ---
                With TDetailsReferencesClient(MemNumLigne)
                    If TexteSansMasque = "" Or .ReferenceClient = "" Then
                        .NbrPieces = 0
                    Else
                        If IsNumeric(TexteSansMasque) = True Then
                            .NbrPieces = CDbl(TexteSansMasque)
                        Else
                             .NbrPieces = 0
                        End If
                    End If
                End With
                
            Case Else
    
        End Select
    
    End If
    
    '--- rendre le contrôle texte invisible ---
    MEBEditionDetailsReferencesClient.Visible = False

    '--- construction de la grille ---
    GestionDetailsReferencesClient GG_AFFICHAGE

End Sub

Private Sub MEBEditionPrevisionnel_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    If InterdireEvenements = False Then
            
        '--- rendre invisible la liste du choix du poste d'anodisation ---
        PBChoixPosteAnodisationPrevisionnel.Visible = False

        With MSHFGPrevisionnel
    
            Select Case .Col
        
                Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                    '--- n° de commande interne ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                Case COLONNES_PREVISIONNEL.C_NBR_REPARATIONS
                    '--- nombre de réparations ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                    
                Case COLONNES_PREVISIONNEL.C_NBR_PIECES
                    '--- nombre de pièces ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                Case COLONNES_PREVISIONNEL.C_NUM_BARRE
                    '--- numéro de barre ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                Case COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                    '--- numéro de la gamme d'anodisation ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                    '--- choix du poste d'anodisation ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                    
                Case Else
        
            End Select
    
        End With

    End If

End Sub

Private Sub MEBEditionPrevisionnel_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    MemNumLigne = 0
    MemNumColonne = 0
    
    '--- rendre visible le focus de la table ---
    SFocusTablePrevisionnel.Visible = True

End Sub

Private Sub MEBEditionPrevisionnel_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With MSHFGPrevisionnel

        '--- analyse de la touche ---
        Select Case KeyCode

            Case vbKeyDown
                '--- flèche basse ---
                .SetFocus
                If .Row < .Rows - 1 Then .Row = .Row + 1
                KeyCode = 0
            
            Case vbKeyUp
                '--- flèche haute ---
                .SetFocus
                If .Row > .FixedRows Then .Row = .Row - 1
                KeyCode = 0
  
            Case Else
  
        End Select
  
    End With

End Sub

Private Sub MEBEditionPrevisionnel_KeyPress(KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    With MSHFGPrevisionnel

        '--- analyse de la touche ---
        Select Case KeyAscii

            Case vbKeyReturn
                '--- touche entrée ---
                Select Case .Col
            
                    Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                        '--- n° de la commande interne ---
                        .Col = COLONNES_PREVISIONNEL.C_NBR_PIECES
                    
                    Case COLONNES_PREVISIONNEL.C_NBR_REPARATIONS
                        '--- nombre de réparations ---
                        .Col = COLONNES_PREVISIONNEL.C_NBR_PIECES
                
                    Case COLONNES_PREVISIONNEL.C_NBR_PIECES
                        '--- nombre de pièces ---
                        .Col = COLONNES_PREVISIONNEL.C_NUM_BARRE
                    
                    Case COLONNES_PREVISIONNEL.C_NUM_BARRE
                        '--- numéro de barre ---
                        .Col = COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                        
                    Case COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                        '--- numéro de gamme d'anodisation ---
                        .Col = COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                        
                    Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                        '--- choix du poste d'anodisation ---
                        If .Row < .Rows - 1 Then .Row = .Row + 1
                        .Col = COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                
                    Case Else

                End Select

                '--- mettre le focus sur le tableau ---
                .SetFocus
                KeyAscii = 0

            Case Else
                Select Case .Col
                    Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES, 8
                    Case COLONNES_PREVISIONNEL.C_NBR_REPARATIONS: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 1
                    Case COLONNES_PREVISIONNEL.C_NBR_PIECES: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6
                    Case COLONNES_PREVISIONNEL.C_NUM_BARRE: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 2
                    Case COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6
                    Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 1
                    Case Else
                End Select

        End Select

    End With

End Sub

Private Sub MEBEditionPrevisionnel_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim ChoixPosteAnodisation As CHOIX_POSTE_ANODISATION                        'choix du poste d'anodisation fonction de l'énumération
    Dim NumGammeAnodisation As String                                                              'numéro d'une gamme d'anodisation
    Dim TexteComplet As String, _
            TexteSansMasque As String
    
    '--- analyse des données frappées ---
    If MemNumLigne > 0 And MemNumColonne > 0 Then
    
        '--- affectation ---
        With MEBEditionPrevisionnel
            TexteComplet = .Text
            TexteSansMasque = .ClipText
        End With
      
        Select Case MemNumColonne
        
            Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                '--- n° de commande interne ---
                If TexteSansMasque = "" Then
                    
                    '--- pas de n° de commande interne ---
                    With TPrevisionnel(MemNumLigne)
                        .NumCommandeInterne = 0
                    End With
                
                Else
                
                    '--- possibilité de valider plusieurs fois le m^me numéro de commande interne ---
                    If InsertionCommandeInterne(TG_PREVISIONNEL, MemNumLigne, TexteSansMasque) <> TROUVE Then
                        
                        '--- message d'erreur ---
                        MessageErreur TITRE_MESSAGES, MESSAGE_122
                        
                        '--- affectation ---
                        With TPrevisionnel(MemNumLigne)
                            .NumCommandeInterne = 0
                        End With
                        
                        '--- replacer le focus sur la grille au bon endroit ---
                        With MSHFGPrevisionnel
                            .Row = MemNumLigne
                            .Col = MemNumColonne
                            .SetFocus
                        End With
                    
                    End If
                
                End If
            
            Case COLONNES_PREVISIONNEL.C_NBR_REPARATIONS
                '--- nombre de réparations ---
                If IsNumeric(TexteSansMasque) = True Then
                    TPrevisionnel(MemNumLigne).NbrReparations = TexteSansMasque
                Else
                    TPrevisionnel(MemNumLigne).NbrReparations = ""
                End If
            
            Case COLONNES_PREVISIONNEL.C_NBR_PIECES
                '--- nombre de pièces ---
                If IsNumeric(TexteSansMasque) = True Then
                    TPrevisionnel(MemNumLigne).NbrPieces = CLng(TexteSansMasque)
                Else
                    TPrevisionnel(MemNumLigne).NbrPieces = 0
                End If
            
            Case COLONNES_PREVISIONNEL.C_NUM_BARRE
                '--- numéro de barre ---
                If IsNumeric(TexteSansMasque) = True Then
                    TPrevisionnel(MemNumLigne).NumBarre = CLng(TexteSansMasque)
                Else
                    TPrevisionnel(MemNumLigne).NumBarre = 0
                End If
            
            Case COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                '--- numéro de la gamme d'anodisation ---
                If IsNumeric(TexteSansMasque) = True Then
                    
                    '--- affectation du numéro de gamme d'anodisation ---
                    NumGammeAnodisation = Right(String(6, "0") & TexteSansMasque, 6)
                    
                    '--- contrôle de l'existence de la gamme ---
                    If ExistenceGammesAnodisation(NumGammeAnodisation) = TROUVE Then
                        
                        '--- affectation du numéro de gamme d'anodisation dans le prévisionnel ---
                        TPrevisionnel(MemNumLigne).NumGammeAnodisation = NumGammeAnodisation
                        
                        '--- chargement de la gamme d'anodisation dans le prévisionnel ---
                        If RechercheGammesAnodisation(NumGammeAnodisation) = TROUVE Then
                            TPrevisionnel(MemNumLigne).TGammesAnodisation = TTempEnrGammesAnodisation
                        End If
                    
                    Else
                        
                        '--- lancement du message d'erreur ---
                        Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_131)
                        MSHFGPrevisionnel.Col = COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                    
                    End If
                
                End If
            
            Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                '--- choix du poste d'anodisation ---
                If IsNumeric(TexteSansMasque) = True Then
                    ChoixPosteAnodisation = CLng(TexteSansMasque)
                    If ChoixPosteAnodisation >= CHOIX_POSTE_ANODISATION.C_AUTOMATIQUE And ChoixPosteAnodisation <= CHOIX_POSTE_ANODISATION.C_C16_IMPOSE Then
                        TPrevisionnel(MemNumLigne).ChoixPosteAnodisation = ChoixPosteAnodisation
                    End If
                End If
            
            Case Else
    
        End Select

    End If
    
    '--- focus ---
    SFocusTablePrevisionnel.Visible = False
    
    '--- rendre le contrôle texte invisible ---
    MEBEditionPrevisionnel.Visible = False

    '--- construction de la grille ---
    GestionPrevisionnel GG_AFFICHAGE

End Sub

Private Sub MEBTempsPhases_Change(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- calcul du temps total du cycle du redresseur ---
    If InterdireEvenements = False Then
        LTempsTotalGammeRedresseur.Caption = Right(CTemps2(CalculTempsTotalGammeRedresseur()), 7)
    End If

End Sub

Private Sub MEBTempsPhases_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With MEBTempsPhases(Index)
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub MEBTempsPhases_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub MEBTempsPhases_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
End Sub

Private Sub MEBTempsPhases_ValidationError(Index As Integer, InvalidText As String, StartPosition As Integer)
    On Error Resume Next
    MEBTempsPhases(Index).Text = Replace(InvalidText, "_", "0")
End Sub

Private Sub MEBTempsReelAnodisation_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With MEBTempsReelAnodisation
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub MEBTempsReelAnodisation_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub MEBTempsReelAnodisation_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
End Sub

Private Sub MEBTempsReelAnodisation_ValidationError(InvalidText As String, StartPosition As Integer)
    On Error Resume Next
    MEBTempsReelAnodisation.Text = Replace(InvalidText, "_", "0")
End Sub

Private Sub MSHFGDetailsCharges_DblClick()
    On Error Resume Next
    InterdireEvenements = True
    EditionChargement vbKeySpace  'simule un espace
    InterdireEvenements = False
End Sub

Private Sub MSHFGDetailsCharges_GotFocus()
    On Error Resume Next
    SFocusTableDetailsCharges.Visible = True
    UtilisationOutilsGrilles TG_CHARGEMENT, True
End Sub

Private Sub MSHFGDetailsCharges_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete: EditionChargement vbKeyBack  'simule un retour arrière (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGDetailsCharges_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionChargement KeyAscii  'envoi de la touche frappée
End Sub

Private Sub MSHFGDetailsCharges_LeaveCell()
    On Error Resume Next
    MEBEditionDetailsCharges.Visible = False
    PBReferencesClient.Visible = False
End Sub

Private Sub MSHFGDetailsCharges_LostFocus()
    On Error Resume Next
    SFocusTableDetailsCharges.Visible = False
    UtilisationOutilsGrilles TG_CHARGEMENT, False
End Sub

Private Sub MSHFGDetailsCharges_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumLigne As Integer, _
           NumColonne As Integer
    Dim LongueurImageCellule As Long, _
            HauteurImageCellule As Long, _
            XInferieurImageCellule As Long, _
            YInferieurImageCellule As Long, _
            XSuperieurImageCellule As Long, _
            YSuperieurImageCellule As Long
    
    With MSHFGDetailsCharges
    
        '--- affectation ---
        NumLigne = .MouseRow
        NumColonne = .MouseCol
           
        '--- analyse en fonction du numéro de ligne ---
        If NumLigne > 0 Then
        
            Select Case NumColonne
            
                Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                    '--- n° de commande interne ---
                    ' il faut détecter le click dans l'image de la cellule
                    If .CellPicture <> LoadPicture() Then
                    
                        '--- recherche des dimensions de l'image se trouvant dans la cellule ---
                        LongueurImageCellule = ILImagesPourGrilles.ImageWidth
                        HauteurImageCellule = ILImagesPourGrilles.ImageHeight
                       
                        '--- coordonnées de l'image ---
                        XInferieurImageCellule = .CellLeft + .CellWidth - (LongueurImageCellule * Screen.TwipsPerPixelX)
                        YInferieurImageCellule = .CellTop + .CellHeight - (HauteurImageCellule * Screen.TwipsPerPixelY)
                        XSuperieurImageCellule = .CellLeft + .CellWidth
                        YSuperieurImageCellule = .CellTop + .CellHeight
                       
                        If X >= XInferieurImageCellule And Y >= YInferieurImageCellule And _
                           X <= XSuperieurImageCellule And Y <= YSuperieurImageCellule Then
                           
                            With PBReferencesClient
                                
                                '--- affiche le contrôle liste au bon endroit (en dessous de la cellule) ---
                                .Move MSHFGDetailsCharges.Left + MSHFGDetailsCharges.CellLeft, _
                                           MSHFGDetailsCharges.Top + MSHFGDetailsCharges.CellTop + MSHFGDetailsCharges.CellHeight
                    
                                '--- mémorisation des valeurs de ligne et colonne ---
                                MemNumLigneDetailsCharges = NumLigne
                                MemNumColonneDetailsCharges = NumColonne
                                
                                '--- affichage de la liste ---
                                GestionDetailsReferencesClient GG_INITIALISATION
                                GestionDetailsReferencesClient GG_TRANSFERT_DONNEES
                                GestionDetailsReferencesClient GG_AFFICHAGE

                                .Visible = True
                                .Refresh
                                MSHFGDetailsReferencesClient.SetFocus
                            
                            End With

                        End If
                    
                    End If
    
                Case Else
        
            End Select
        
        End If
    
    End With

End Sub

Private Sub MSHFGDetailsCharges_Scroll()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rendre invisible le champ d'édition en cas de scrolling ---
    If MEBEditionDetailsCharges.Visible = True Then
        MEBEditionDetailsCharges.Visible = False
    End If
    
    '--- rendre invisible la liste des références du client en cas de scrolling ---
    If PBReferencesClient.Visible = True Then
        PBReferencesClient.Visible = False
    End If

End Sub

Private Sub MSHFGDetailsReferencesClient_DblClick()
    On Error Resume Next
    InterdireEvenements = True
    EditionDetailsReferencesClient vbKeySpace  'simule un espace
    InterdireEvenements = False
End Sub

Private Sub MSHFGDetailsReferencesClient_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete: EditionDetailsReferencesClient vbKeyBack  'simule un retour arrière (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGDetailsReferencesClient_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionDetailsReferencesClient KeyAscii  'envoi de la touche frappée
End Sub

Private Sub MSHFGDetailsReferencesClient_LeaveCell()
    On Error Resume Next
    MEBEditionDetailsReferencesClient.Visible = False
End Sub

Private Sub MSHFGDetailsReferencesClient_Scroll()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rendre invisible le champ d'édition en cas de scrolling ---
    If MEBEditionDetailsReferencesClient.Visible = True Then
        MEBEditionDetailsReferencesClient.Visible = False
    End If

End Sub

Private Sub MSHFGPrevisionnel_DblClick()
    On Error Resume Next
    InterdireEvenements = True
    EditionPrevisionnel vbKeySpace  'simule un espace
    InterdireEvenements = False
End Sub

Private Sub MSHFGPrevisionnel_GotFocus()
    On Error Resume Next
    SFocusTablePrevisionnel.Visible = True
    UtilisationOutilsGrilles TG_PREVISIONNEL, True
End Sub

Private Sub MSHFGPrevisionnel_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete: EditionPrevisionnel vbKeyBack  'simule un retour arrière (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGPrevisionnel_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionPrevisionnel KeyAscii  'envoi de la touche frappée
End Sub

Private Sub MSHFGPrevisionnel_LeaveCell()
    On Error Resume Next
    MEBEditionPrevisionnel.Visible = False
    PBChoixPosteAnodisationPrevisionnel.Visible = False
End Sub

Private Sub MSHFGPrevisionnel_LostFocus()
    On Error Resume Next
    SFocusTablePrevisionnel.Visible = False
    UtilisationOutilsGrilles TG_PREVISIONNEL, False
End Sub

Private Sub MSHFGPrevisionnel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumLigne As Integer, _
           NumColonne As Integer
    Dim LongueurImageCellule As Long, _
            HauteurImageCellule As Long, _
            XInferieurImageCellule As Long, _
            YInferieurImageCellule As Long, _
            XSuperieurImageCellule As Long, _
            YSuperieurImageCellule As Long
    
    With MSHFGPrevisionnel
    
        '--- affectation ---
        NumLigne = .MouseRow
        NumColonne = .MouseCol
           
        '--- analyse en fonction du numéro de ligne ---
        If NumLigne > 0 Then
        
            Select Case NumColonne
            
                Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                    '--- choix du poste d'anodisation ---
                    ' il faut détecter le click dans l'image de la cellule
                    If .CellPicture <> LoadPicture() Then
                    
                        '--- recherche des dimensions de l'image se trouvant dans la cellule ---
                        LongueurImageCellule = ILImagesPourGrilles.ImageWidth
                        HauteurImageCellule = ILImagesPourGrilles.ImageHeight
                       
                        '--- coordonnées de l'image ---
                        XInferieurImageCellule = .CellLeft + .CellWidth - (LongueurImageCellule * Screen.TwipsPerPixelX)
                        YInferieurImageCellule = .CellTop + .CellHeight - (HauteurImageCellule * Screen.TwipsPerPixelY)
                        XSuperieurImageCellule = .CellLeft + .CellWidth
                        YSuperieurImageCellule = .CellTop + .CellHeight
                       
                        If X >= XInferieurImageCellule And Y >= YInferieurImageCellule And _
                           X <= XSuperieurImageCellule And Y <= YSuperieurImageCellule Then
                           
                            With PBChoixPosteAnodisationPrevisionnel
                                
                                '--- affiche le contrôle liste au bon endroit (en dessous de la cellule) ---
                                .Move MSHFGPrevisionnel.Left + MSHFGPrevisionnel.CellLeft, _
                                           MSHFGPrevisionnel.Top + MSHFGPrevisionnel.CellTop + MSHFGPrevisionnel.CellHeight
                    
                                '--- affichage de la boite de dialogues ---
                                .Visible = True
                                .Refresh
                                MSHFGPrevisionnel.SetFocus
                            
                            End With

                        End If
                    
                    End If
    
                Case Else
        
            End Select
        
        End If
    
    End With

End Sub

Private Sub MSHFGPrevisionnel_Scroll()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rendre invisible le champ d'édition en cas de scrolling ---
    If MEBEditionPrevisionnel.Visible = True Then
        MEBEditionPrevisionnel.Visible = False
    End If
    
    '--- rendre invisible la liste d'édition en cas de scrolling ---
    If PBChoixPosteAnodisationPrevisionnel.Visible = True Then
        PBChoixPosteAnodisationPrevisionnel.Visible = False
    End If

End Sub

Private Sub PBBoutons_Resize()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBReduire.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBReduire.Width
    LDonneesTransmisesAPI.Left = CBReduire.Left - 68 * MARGES.M_ENTRE_BOUTONS - LDonneesTransmisesAPI.Width
    
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
' Rôle      : Gére l'états des boutons après une action de l'opèrateur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionBoutons(ByVal Situation As ETATS_BOUTONS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case Situation
        
        Case ETATS_BOUTONS.E_CHARGEMENT_FENETRE
            '--- au chargement de la fenetre ---
            CBQuitter.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---
            CBQuitter.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
            CBQuitter.Enabled = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            CBQuitter.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- après nouveau ---
            CBQuitter.Enabled = True
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---
            CBQuitter.Enabled = True
        
        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

End Sub

Private Sub PBCharges_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case Source.Name
        
        Case IAutorisationChargement.Name
            '--- dépose d'une source venant de l'autorisation de chargement sur une des charges ---
             With TEtatsPostes(Index)
                If .EtatsChariots = E_PRESENT_VERROUILLE Then
                    IntroductionChargeAuChargement Index
                End If
            End With
        
        Case Else

    End Select

End Sub

Private Sub PBCharges_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    Select Case Source.Name
        
        Case IAutorisationChargement.Name
             '--- la source est l'autorisation de chargement ---
             Select Case State
                 
                 Case 0
                     '--- 0 = Entre (le contrôle source entre dans la portée de la cible) ---
                     With TEtatsPostes(Index)
                        If .EtatsChariots = E_PRESENT_VERROUILLE And .NumCharge = 0 Then
                           Source.DragIcon = ILIcones.ListImages("etoile").ExtractIcon
                        Else
                           Source.DragIcon = ILIcones.ListImages("sens interdit").ExtractIcon
                        End If
                    End With
                 
                 Case 1
                     '--- 1 = Sort (le contrôle source sort de la portée de la cible) ---
                     Source.DragIcon = ILIcones.ListImages("sens interdit").ExtractIcon
            
                 Case 2
                      '--- 2 = Dessus (le contrôle source est passé d'une position à une autre dans la cible) ---
                 
                 Case Else
            
             End Select
                
        Case Else

    End Select

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

Private Sub IEtatsPostes_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumCharge As Integer
    
    If Button = vbLeftButton Then
        
        '****************************************************************************************************************
        '                                     forcer l'entrée de la charge dans la ligne (clic gauche)
        '****************************************************************************************************************
        
        '--- affectation ---
        NumCharge = TEtatsPostes(Index).NumCharge
        
        '--- forcer l'entrée de la charge dans la ligne ---
        If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
            
            With TEtatsCharges(NumCharge)
            
                If .PtrZoneGammeAnodisation = 0 Then
                    
                    If AppelFenetre(F_MESSAGE, _
                                            TITRE_MESSAGES, _
                                            vbCrLf & _
                                            "Cette action FORCERA l'ENTREE de la CHARGE " & NumCharge & vbCrLf & _
                                            "le PLUS TOT POSSIBLE dans la ligne." & vbCrLf & vbCrLf & _
                                            "cs|VOTRE RESPONSABILITE EST ENGAGEE" & vbCrLf & vbCrLf & _
                                            "c|Voulez-vous réellement EFFECTUER cette ACTION ?", _
                                            TYPES_MESSAGES.T_ATTENTION, _
                                            TYPES_BOUTONS.T_OUI_NON, _
                                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                        '--- pointer la première zone ---
                        .PtrZoneGammeAnodisation = 1
    
                    End If
            
                End If
    
            End With
    
        End If
    
    Else
        
        '****************************************************************************************************************
        '                                            condamnation du poste (clic droit de la souris)
        '****************************************************************************************************************
        CondamnationPoste Index, TITRE_MESSAGES
    
    End If

End Sub

Private Sub PBPrevisionnel_Resize()
    On Error Resume Next
    Cadre3DSurImage Me.PBPrevisionnel, 2, True, False
End Sub

Private Sub PBReferencesClient_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    'On Error Resume Next
    
    '--- rendre visible le focus de la table ---
    'SFocusTablePrevisionnel.Visible = True
    
    '--- déplacer le focus ---
    'LBNumLignesReferencesClient.SetFocus

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
' Rôle      : Permet ou non l'utilisation des outils des grilles
' Entrées :  Typegrille -> fonction de l'énumération TYPES_GRILLES
'                 Autorisation -> TRUE = Autorise l'utilisation des outils grilles, FALSE = Verrouille les outils grilles
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub UtilisationOutilsGrilles(ByVal TypeGrille As TYPES_GRILLES, _
                                                          ByVal Autorisation As Boolean)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    '--- autorisation ou non sur les outils concernés ---
    Select Case TypeGrille
    
        Case TYPES_GRILLES.TG_CHARGEMENT
            '--- grille du chargement ---
            TOBGestionGrilleChargement(0).buttons("SupprimerLigne").Enabled = Autorisation
            TOBGestionGrilleChargement(0).buttons("InsererLigne").Enabled = Autorisation
            TOBGestionGrilleChargement(0).buttons("CompacterGrille").Enabled = Autorisation
            
        Case TYPES_GRILLES.TG_PREVISIONNEL
            '--- grille du prévisionnel ---
            TOBGestionGrillePrevisionnel(0).buttons("SupprimerLigne").Enabled = Autorisation
            TOBGestionGrillePrevisionnel(0).buttons("InsererLigne").Enabled = Autorisation
            TOBGestionGrillePrevisionnel(0).buttons("CompacterGrille").Enabled = Autorisation
    
        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées : OngletChoisie -> onglet choisie à l'ouverture en fonction de l'énumération des ONGLETS
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal OngletChoisie As ONGLETS_CHARGEMENT_PREVISIONNEL)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim OCCCroixSelection As CheckBox

    '--- modification de la forme du chargement prévisionnel ---
    Select Case OngletChoisie
        Case ONGLETS_CHARGEMENT_PREVISIONNEL.O_CHARGEMENT, ONGLETS_CHARGEMENT_PREVISIONNEL.O_PREVISIONNEL
            CTOnglets.CurrTab = OngletChoisie
        Case Else
    End Select
    
    '--- initialisation des grilles et affichage ---
    GestionGrilleRecherche O_CHARGEMENT, GG_INITIALISATION
    GestionGrilleRecherche O_CHARGEMENT, GG_AFFICHAGE
    
    GestionGrilleRecherche O_PREVISIONNEL, GG_INITIALISATION
    GestionGrilleRecherche O_PREVISIONNEL, GG_AFFICHAGE
    
    GestionDetailsCharges GG_INITIALISATION
    GestionDetailsCharges GG_AFFICHAGE
    
    GestionPrevisionnel GG_INITIALISATION
    GestionPrevisionnel GG_AFFICHAGE
    
    '--- sélectionner les croix des options des ponts pour la petite vitesse ---
    For Each OCCCroixSelection In CBOptionsPonts
        Select Case OCCCroixSelection.Index
            Case OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV, OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV
                OCCCroixSelection.value = vbChecked
            Case Else
                OCCCroixSelection.value = vbUnchecked
        End Select
    Next
    
    '--- désélectionner l'activation de l'air dans le bain de brillantange ---
    CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).value = vbUnchecked
    
    '--- Initialise les champs de la partie redresseur ---
    InitialisationChampsRedresseur
    
    '--- visualisation des différents états du chargement prévisionnel ---
    EtatsChargementPrevisionnel
    
    '--- calcul du prévisionnel ---
    CalculPrevisionnel
    
    '--- lancement des timers ---
    TimerChargementPrevisionnel.Enabled = True
    TimerCalculPrevisionnel.Enabled = True

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise les champs de la partie redresseur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationChampsRedresseur()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const INIT_TEMPS As String = "0:00:00"
    Const INIT_TENSION As String = "0,0"
    Const INIT_INTENSITE As String = "0"

    '--- déclaration ---
    Dim a As Integer                                            'pour les boucles FOR...NEXT

    '--- affectation ---
    
    '--- interdire les évènements ---
    InterdireEvenements = True

    '--- forçage du mode U ou I en mode tension ---
    Call LModeUouI_Click(MODES_U_OU_I.M_TENSION)

    '--- initialisation des champs temps, tension, intensité ---
    For a = MEBTempsPhases.LBound To MEBTempsPhases.UBound
        MEBTempsPhases(a).Text = INIT_TEMPS
        TBTensionsPhases(a).Text = INIT_TENSION
        TBIntensitesPhases(a).Text = INIT_INTENSITE
    Next a

    '--- temps total de la gamme redresseur ---
    LTempsTotalGammeRedresseur.Caption = INIT_TEMPS

    '--- autorisation des évènements ---
    InterdireEvenements = False
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---

    '--- déclaration ---

    '--- affectation ---
  
    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondOrange1
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- construction des cadres 3D ---
    Cadre3DSurImage Me.PBPrevisionnel, 2, True, False
    'Cadre3DSurImage Me.PBMoteurInference, 2, True, False
    
    '--- valeur par défaut du champ du temps réel d'anodisation ---
    With MEBTempsReelAnodisation
        If .Text = "" Then
            .Text = "00:00:00"
        End If
        .Enabled = False
    End With
    
    '--- affichage du temps de compensation d'anodisation ---
    InterdireEvenements = True
    TBCompensation.Text = TempsCompensationAnodisationMinutes
    CBRechercherPar(ONGLETS_CHARGEMENT_PREVISIONNEL.O_CHARGEMENT).ListIndex = IDX_RECHERCHER_PAR.IDX_NUM_GAMME - 1      '-1 car l'index démarre de 0
    CBRechercherPar(ONGLETS_CHARGEMENT_PREVISIONNEL.O_PREVISIONNEL).ListIndex = IDX_RECHERCHER_PAR.IDX_NUM_GAMME - 1     '-1 car l'index démarre de 0
    InterdireEvenements = False
    Dim a As Integer
    '--
     For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a).DefinitionPoste
            CBNumPosteDepart.AddItem (.NomPoste & " - " & .LibellePoste)
            CBNumPosteDepart.ItemData(CBNumPosteDepart.NewIndex) = .NumPoste

        End With
    Next a
    
    ' TODO add barre 202411
    ComboBarre.AddItem ("")
    ComboBarre.ItemData(ComboBarre.NewIndex) = ""
    For a = LBound(TBarres()) To UBound(TBarres())
        ComboBarre.AddItem (TBarres(a).Libelle)
        ComboBarre.ItemData(ComboBarre.NewIndex) = TBarres(a).Libelle
        
        
    Next a
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE
    
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
    
    '--- neutralisation des timers ---
    With TimerChargementPrevisionnel
        .Enabled = False
        .Interval = 0
    End With
    With TimerCalculPrevisionnel
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFChargementPrevisionnel = Nothing

End Sub

Private Sub TBCommencantPar_GotFocus(Index As Integer)
    On Error Resume Next
    With TBCommencantPar(Index)
        If .SelText = "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Sub TBCommencantPar_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri Index
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBCommencantPar_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    Select Case Succ(CBRechercherPar(Index).ListIndex)
        Case IDX_RECHERCHER_PAR.IDX_NUM_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6                                'n° de gamme
        Case IDX_RECHERCHER_PAR.IDX_REF_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE, 30                                         'référence de la gamme
        Case IDX_RECHERCHER_PAR.IDX_NOM_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE, 50                                        'nom de la gamme
        Case Else
    End Select
End Sub

Private Sub TBCompensation_Change()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    If InterdireEvenements = False Then

        With TBCompensation
            If IsNumeric(.Text) = True Then
                TempsCompensationAnodisationMinutes = CInt(.Text)
            Else
                TempsCompensationAnodisationMinutes = 0
            End If
        End With
    
    End If

End Sub

Private Sub TBCompensation_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBCompensation
        If IsNumeric(.Text) = True Then
            .Text = CStr(CInt(.Text))
        Else
            .Text = ""
        End If
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub TBCompensation_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBCompensation_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_RELATIFS, 3
End Sub

Private Sub TBCompensation_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBCompensation
        If IsNumeric(.Text) = True Then
            .Text = Format(CLng(.Text), FORMAT_COMPENSATION)
        Else
            .Text = "0"
        End If
    End With

End Sub

Private Sub TBContenant_GotFocus(Index As Integer)
    On Error Resume Next
    With TBContenant(Index)
        If .SelText = "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Sub TBContenant_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri Index
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBContenant_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES
End Sub

Private Sub TBDelaiSupStabilisationCharge_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBDelaiSupStabilisationCharge
        If IsNumeric(.Text) = True Then
            .Text = CStr(CLng(.Text))
        Else
            .Text = ""
        End If
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TBDelaiSupStabilisationCharge_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBDelaiSupStabilisationCharge_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
End Sub

Private Sub TBDelaiSupStabilisationCharge_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBDelaiSupStabilisationCharge
        If IsNumeric(.Text) = True Then
            .Text = Format(CLng(.Text), FORMAT_DELAI_SUP_STABILISATION_CHARGE)
        Else
            .Text = ""
        End If
    End With

End Sub

Private Sub TBIntensitesPhases_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim Intensite As Integer
    
    With TBIntensitesPhases(Index)
        If IsNumeric(.Text) = True Then
            Intensite = CInt(.Text)
            If Intensite > LIMITES_REDRESSEURS.LM_INTENSITE Then
                Intensite = LIMITES_REDRESSEURS.LM_INTENSITE
            End If
            .Text = Format(Intensite, FORMAT_INTENSITE_ENTIER)
        Else
            .Text = "0"
        End If
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With
    
End Sub

Private Sub TBIntensitesPhases_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBIntensitesPhases_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 4
End Sub

Private Sub TBIntensitesPhases_LostFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim Intensite As Single
    
    With TBIntensitesPhases(Index)
        If IsNumeric(.Text) = True Then
            Intensite = CInt(.Text)
            If Intensite > LIMITES_REDRESSEURS.LM_INTENSITE Then
                Intensite = LIMITES_REDRESSEURS.LM_INTENSITE
            End If
            .Text = Format(Intensite, FORMAT_INTENSITE_ENTIER)
        Else
            .Text = "0"
        End If
    End With

End Sub

Private Sub TBNumGammeAnodisation_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBNumGammeAnodisation
        If IsNumeric(.Text) = True Then
            .Text = CStr(CLng(.Text))
        Else
            .Text = ""
        End If
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TBNumGammeAnodisation_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBNumGammeAnodisation_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS
End Sub

Private Sub TBNumGammeAnodisation_KeyUp(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim NumGammeTexte  As String
    
    '--- chargement de la gamme ---
    Select Case KeyCode
        Case vbKeyDown, vbKeyUp, vbKeyTab, vbKeyShift
        Case Else
            With TBNumGammeAnodisation
                If IsNumeric(.Text) = True Then
                    NumGammeTexte = Format(CLng(.Text), FORMAT_NUM_GAMME_ANODISATION)
                    ChargeGammeAnodisationChargement NumGammeTexte
                Else
                    ChargeGammeAnodisationChargement FORMAT_NUM_GAMME_ANODISATION
                End If
            End With
    End Select

End Sub

Private Sub TBNumGammeAnodisation_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    With TBNumGammeAnodisation
        If IsNumeric(.Text) = True Then
            .Text = Format(CLng(.Text), FORMAT_NUM_GAMME_ANODISATION)
        Else
            .Text = FORMAT_NUM_GAMME_ANODISATION
        End If
    End With

End Sub

Private Sub TBTensionsPhases_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim Tension As Single
    
    With TBTensionsPhases(Index)
        If IsNumeric(.Text) = True Then
            Tension = CSng(.Text)
            If Tension > LIMITES_REDRESSEURS.LM_TENSION Then
                Tension = LIMITES_REDRESSEURS.LM_TENSION
            End If
            .Text = Format(Tension, FORMAT_TENSION_1_DECIMALE)
        Else
            .Text = "0.0"
        End If
        .SelStart = 0          'met en surbrillance la sélection saisie
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub TBTensionsPhases_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBTensionsPhases_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_REELS_POSITIFS, 4
End Sub

Private Sub TBTensionsPhases_LostFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim Tension As Single
    
    With TBTensionsPhases(Index)
        If IsNumeric(.Text) = True Then
            Tension = CSng(.Text)
            If Tension > LIMITES_REDRESSEURS.LM_TENSION Then
                Tension = LIMITES_REDRESSEURS.LM_TENSION
            End If
            .Text = Format(Tension, FORMAT_TENSION_1_DECIMALE)
        Else
            .Text = "0.0"
        End If
    End With

End Sub

Private Sub TDBGGrilleRecherche_DblClick(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim NumLigne As Integer
    
    Dim NumGammeTexte  As String
    
    '--- affichage de la gamme ---
    If Index = ONGLETS_CHARGEMENT_PREVISIONNEL.O_CHARGEMENT Then
    
        '--- onglet du chargement ---
        With TBNumGammeAnodisation
            .Text = ADODCGammesAnodisation(Index).Recordset.Fields("NumGamme")
            If IsNumeric(.Text) = True Then
                NumGammeTexte = Format(CLng(.Text), FORMAT_NUM_GAMME_ANODISATION)
                ChargeGammeAnodisationChargement NumGammeTexte
            Else
                ChargeGammeAnodisationChargement FORMAT_NUM_GAMME_ANODISATION
            End If
        End With

    Else

        '--- onglet du prévisionnel ---
        NumLigne = MSHFGPrevisionnel.Row
        
        If NumLigne >= 1 And NumLigne <= NBR_LIGNES_PREVISIONNEL Then
        
            With TPrevisionnel(NumLigne)
            
                If .NumCommandeInterne > 0 Then
                
                    '--- affectation du numéro de gamme ---
                    .NumGammeAnodisation = ADODCGammesAnodisation(Index).Recordset.Fields("NumGamme")
                    .ChoixPosteAnodisation = C_AUTOMATIQUE
                
                End If
        
                '--- rafraichir la grille ---
                GestionPrevisionnel GG_AFFICHAGE
            
            End With
            
        End If
            
    End If

End Sub

Private Sub TDBGGrilleRecherche_Error(Index As Integer, ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = vbDataErrContinue
End Sub

Private Sub TDBGGrilleRecherche_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim NumGammeTexte  As String
    
    '--- appel de la routine ---
    Select Case KeyCode
        
        Case vbKeyReturn
            '--- affichage de la gamme ---
            With TBNumGammeAnodisation
                .Text = ADODCGammesAnodisation(Index).Recordset.Fields("NumGamme")
                If IsNumeric(.Text) = True Then
                    NumGammeTexte = Format(CLng(.Text), FORMAT_NUM_GAMME_ANODISATION)
                    ChargeGammeAnodisationChargement NumGammeTexte
                Else
                    ChargeGammeAnodisationChargement FORMAT_NUM_GAMME_ANODISATION
                End If
            End With
            KeyCode = 0: Shift = 0
        
        Case vbKeyHome
            If Shift = vbCtrlMask Then
                ADODCGammesAnodisation(Index).Recordset.MoveFirst
                KeyCode = 0: Shift = 0
            End If
        
        Case vbKeyEnd
            If Shift = vbCtrlMask Then
                ADODCGammesAnodisation(Index).Recordset.MoveLast
                KeyCode = 0: Shift = 0
            End If
        
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown
        
        Case Else
            KeyCode = 0: Shift = 0
    
    End Select

End Sub

Private Sub TimerCalculPrevisionnel_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine ---
    TimerCalculPrevisionnel.Enabled = False
    CalculPrevisionnelAvecAffichageChoix
    TimerCalculPrevisionnel.Enabled = True
    
End Sub

Private Sub TimerChargementPrevisionnel_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine ---
    TimerChargementPrevisionnel.Enabled = False
    EtatsChargementPrevisionnel
    TimerChargementPrevisionnel.Enabled = True
    
End Sub

Private Sub TOBGestionGrilleChargement_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key

        Case "EffacerGrille"

            '--- effacement de la table du chargement ---
            If AppelFenetre(F_MESSAGE, _
                                    TITRE_MESSAGES, _
                                    MESSAGE_300, _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
        
                '--- effacement du tableau commun ---
                Erase TChargement.TDetailsCharges()
        
                '--- réaffichage ---
                GestionDetailsCharges GG_INITIALISATION
                GestionDetailsCharges GG_AFFICHAGE
            
            End If
        
        Case "SupprimerLigne"
            '--- supprime une ligne sur une grille ---
            SupprimerLigneGrille TG_CHARGEMENT
            
        Case "CompacterGrille"
            '--- compacte une grille ---
            CompacterGrille TG_CHARGEMENT
            
        Case "InsererLigne"
            '--- insère une ligne dans une grille ---
            InsererLigneGrille TG_CHARGEMENT
            
        Case Else
    End Select

End Sub

Private Sub TOBGestionGrillePrevisionnel_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key

        Case "EffacerGrille"

            '--- effacement de la table du prévisionnel ---
            If AppelFenetre(F_MESSAGE, _
                                    TITRE_MESSAGES, _
                                    MESSAGE_301, _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
        
                '--- effacement du tableau commun ---
                Erase TPrevisionnel()
        
                '--- réaffichage ---
                GestionPrevisionnel GG_INITIALISATION
                GestionPrevisionnel GG_AFFICHAGE
            
            End If
        
        Case "SupprimerLigne"
            '--- supprime une ligne sur une grille ---
            SupprimerLigneGrille TG_PREVISIONNEL
        
        Case "CompacterGrille"
            '--- compacte une grille ---
            CompacterGrille TG_PREVISIONNEL
        
        Case "InsererLigne"
            '--- insère une ligne dans une grille ---
            InsererLigneGrille TG_PREVISIONNEL

        Case Else
    End Select

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Construction d'un cadre en 3D pour l'esthétique d'un image
' Entrées : ObjetPictureBox -> Objet PictureBox ou l'on doit tracer le cadre 3D
'                      Epaisseur3D -> Détermine le nombre de lignes composant l'effet 3D
'                              Bordure -> TRUE = Tracer d'un rectangle de contour, FALSE = Sans rectangle de contour
'                    CoinsArrondis -> TRUE = Tracer de coins arrondis, FALSE = Sans coins arrondis
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Cadre3DSurImage(ByRef ObjetPictureBox As PictureBox, _
                                                  ByVal Epaisseur3D As Integer, _
                                                  ByVal Bordure As Boolean, _
                                                  ByVal CoinsArrondis As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
        
    '--- déclaration ---
    Dim a As Integer
    Dim CouleurConteneurImage As Long
    Dim TwipsParPixelX As Single, _
            TwipsParPixelY As Single, _
            XInferieurRectangle As Single, _
            YInferieurRectangle As Single, _
            XSuperieurRectangle As Single, _
            YSuperieurRectangle As Single
    
    '--- affectation ---
    TwipsParPixelX = Screen.TwipsPerPixelX
    TwipsParPixelY = Screen.TwipsPerPixelY
    
    With ObjetPictureBox

        .AutoRedraw = True

        .Cls

        '--- tracer de la bordure ---
        If Bordure = True Then
            
            '--- affectation ---
            XInferieurRectangle = .ScaleLeft
            YInferieurRectangle = .ScaleTop
            XSuperieurRectangle = .ScaleWidth - TwipsParPixelX
            YSuperieurRectangle = .ScaleHeight - TwipsParPixelY
            
            '--- tracer ---
            ObjetPictureBox.Line (XInferieurRectangle, YInferieurRectangle)-(XSuperieurRectangle, YSuperieurRectangle), COULEURS.NOIR, B
        
        End If

        '--- tracer de la forme 3D ---
        For a = 1 - Abs(Not (Bordure)) To Epaisseur3D - Abs(Not (Bordure))

            '--- affectation ---
            XInferieurRectangle = .ScaleLeft + a * TwipsParPixelX
            YInferieurRectangle = .ScaleTop + a * TwipsParPixelY
            XSuperieurRectangle = .ScaleWidth - TwipsParPixelX - a * TwipsParPixelX
            YSuperieurRectangle = .ScaleHeight - TwipsParPixelY - a * TwipsParPixelY
            
            '--- construction du rectangle principal ---
            ObjetPictureBox.Line (XInferieurRectangle, YInferieurRectangle)-(XSuperieurRectangle, YSuperieurRectangle), COULEURS.BLANC, B
            
            '--- construction de l'effet 3D ---
            ObjetPictureBox.Line (XSuperieurRectangle, a * TwipsParPixelX)-(XSuperieurRectangle, YSuperieurRectangle), COULEURS.GRIS_3
            ObjetPictureBox.Line -(XInferieurRectangle - TwipsParPixelX, YSuperieurRectangle), COULEURS.GRIS_3
            
        Next a
        
        '--- tracer des coins arrondis ---
        If CoinsArrondis = True Then
            
            '--- affectation ---
            XInferieurRectangle = .ScaleLeft
            YInferieurRectangle = .ScaleTop
            XSuperieurRectangle = .ScaleWidth - TwipsParPixelX
            YSuperieurRectangle = .ScaleHeight - TwipsParPixelY
            
            '--- couleur du conteneur de l'image ---
            CouleurConteneurImage = ObjetPictureBox.Container.BackColor
            
            '--- tracer ---
            ObjetPictureBox.PSet (XInferieurRectangle, YInferieurRectangle), CouleurConteneurImage
            ObjetPictureBox.PSet (XSuperieurRectangle, YInferieurRectangle), CouleurConteneurImage
            ObjetPictureBox.PSet (XSuperieurRectangle, YSuperieurRectangle), CouleurConteneurImage
            ObjetPictureBox.PSet (XInferieurRectangle, YSuperieurRectangle), CouleurConteneurImage
        
        End If
        
        .AutoRedraw = False

    End With
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effacement complet du chargement avec un rafraichissement de l'écran
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EffacementCompletChargement()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim FicheVideChargement As VarChargement
    Dim OCCCroixSelection As CheckBox
    
    '--- affectation ---
    TChargement = FicheVideChargement
    
    '--- rafraichissement des détails ---
    GestionDetailsCharges GG_AFFICHAGE
    
    '--- effacement de la partie gamme ---
    TBNumGammeAnodisation.Text = ""
    LNomGammeAnodisation.Caption = ""
    With MEBTempsReelAnodisation
        .Text = "00:00:00"
        .Enabled = False
    End With
    LTempsAnodisationGamme.Caption = ""
    CBChoixPosteAnodisation.ListIndex = -1
    ComboBarre.ListIndex = 0
    '--- rendre invisible la partie du poste d'anodisation ---
    FTempsEtPosteAnodisation.Visible = False

    '--- effacement de la partie des redresseurs ---
    FRedresseurs.Visible = False
    
    '--- initialise les champs de la partie redresseur ---
    InitialisationChampsRedresseur
    
    '--- effacement des options ---
    FOptions.Visible = False
    
    '--- effacement du numéro de barre ---
    Call ComboBarre_Click
    FNumBarres.Visible = False
    RepositonnerCadre.Visible = False
    
    '--- sélectionner les croix des options des ponts pour la petite vitesse ---
    For Each OCCCroixSelection In CBOptionsPonts
        Select Case OCCCroixSelection.Index
            Case OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV, OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV
                OCCCroixSelection.value = vbChecked
            Case Else
                OCCCroixSelection.value = vbUnchecked
        End Select
    Next

    '--- désélectionner l'activation de l'air dans le bain de brillantange ---
    CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).value = vbUnchecked

    '--- effacer le délai supplémentaire de stabilisation de la charge ---
    TBDelaiSupStabilisationCharge.Text = ""
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lance le calcul du prévisionnel avec l'affichage des choix
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CalculPrevisionnelAvecAffichageChoix()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    If SFocusTablePrevisionnel.Visible = False Or BoutonCalculerALeFocus = True Then
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- changement de la couleur de fond du bouton ---
        With CBCalculerPrevisionnel
            .BackColor = COULEURS.ROUGE_0
            .Refresh
        End With
    
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- calcul du prévisionnel ---
        CalculPrevisionnel
            
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- affichage du choix du prévisionnel ---
        GestionPrevisionnel GG_COMPRESSION
        GestionPrevisionnel GG_AFFICHAGE
        
        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        '--- changement de la couleur de fond du bouton ---
        With CBCalculerPrevisionnel
            .BackColor = COULEURS.BLANC
            .Refresh
        End With
        
    End If
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états du chargement prévisionnel
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsChargementPrevisionnel()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const IDX_CHARIOT_PRESENT As Integer = 58
    Const IDX_CHARIOT_PRESENT_VERROUILLE = 59
    Const AJOUT_IMG_POSTE_CHGT_3_ET_4 As Integer = 60
    Const RECTANGLE_VERT As String = "rectangle vert"
    Const CROIX_DE_CONDAMNATION As String = "croix de condamnation"
    
    '--- déclaration ---
    Dim a As Integer
    Dim NumImage As Integer

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- états des postes ---
    For a = POSTES.P_CHGT_1 To POSTES.P_CHGT_2
    
        With TEtatsPostes(a)
            
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- indication de la condamnation ---
            If .Condamnation = True Then
                
                '--- affichage de la croix de condamnation ---
                If IEtatsPostes(a).Picture <> ILOutilsDivers.ListImages(CROIX_DE_CONDAMNATION).Picture Then
                    Set IEtatsPostes(a).Picture = ILOutilsDivers.ListImages(CROIX_DE_CONDAMNATION).Picture
                End If
                
            Else
                    
                '--- suppression de la croix de condamnation si nécessaire ---
                If IEtatsPostes(a).Picture <> ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture Then
                    Set IEtatsPostes(a).Picture = ILOutilsDivers.ListImages(RECTANGLE_VERT).Picture
                End If

            End If
    
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
            '--- indication des chariots et charges ---
            Select Case .EtatsChariots
            
                Case ETATS_CHARIOTS.E_ABSENT
                    '--- chariot absent ---
                    If PBCharges(a).Picture <> LoadPicture() Then
                        Set PBCharges(a).Picture = LoadPicture()
                        LNomsPostes(a).BackColor = COULEURS.BLANC
                        LNomsPostes(a).ForeColor = COULEURS.NOIR
                    End If
                
                Case ETATS_CHARIOTS.E_PRESENT
                    '--- chariot présent (non verrouillé) ---
                    If PBCharges(a).Picture <> PCCharges.GraphicCell(IDX_CHARIOT_PRESENT) Then
                         Set PBCharges(a).Picture = PCCharges.GraphicCell(IDX_CHARIOT_PRESENT)
                         LNomsPostes(a).BackColor = COULEURS.BLANC
                         LNomsPostes(a).ForeColor = COULEURS.NOIR
                    End If
            
                Case ETATS_CHARIOTS.E_PRESENT_VERROUILLE
                    '--- chariot présent et verrouillé (avec charge ou non) ---
                    If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                        
                        '--- affectation du numéro de l'image ---
                        If ModeAffichageSynoptique = MA_NUM_BARRES Then
                            NumImage = TEtatsCharges(.NumCharge).NumBarre
                        Else
                            NumImage = .NumCharge
                        End If
                        
                        '--- chariot verrouillé avec une charge ---
                        If a = POSTES.P_CHGT_1 Or a = POSTES.P_CHGT_2 Then                              'changement de l'image pour les postes 3 et 4
                            NumImage = NumImage + AJOUT_IMG_POSTE_CHGT_3_ET_4
                        End If
                        
                        '--- choix de l'affichage entre les numéros de charges et les numéros de barres ---
                        If ModeAffichageSynoptique = MA_NUM_CHARGES Then
                            
                            '--- affichage par les numéros de charges ---
                            If PBCharges(a).Picture <> PCCharges.GraphicCell(NumImage) Then
                                 Set PBCharges(a).Picture = PCCharges.GraphicCell(NumImage)
                            End If
                        
                        Else
                            
                            '--- affichage par les numéros de barres ---
                            If PBCharges(a).Picture <> PCBarres.GraphicCell(NumImage) Then
                                 Set PBCharges(a).Picture = PCBarres.GraphicCell(NumImage)
                            End If
                        
                        End If
                        
                        '--- charge prioritaire ---
                        If TEtatsCharges(.NumCharge).ChargePrioritaire = True Then
                            If LNomsPostes(a).BackColor <> COULEURS.BLEU_3 Then
                                LNomsPostes(a).BackColor = COULEURS.BLEU_3
                                LNomsPostes(a).ForeColor = COULEURS.JAUNE_3
                            End If
                        Else
                            If LNomsPostes(a).BackColor <> COULEURS.BLANC Then
                                LNomsPostes(a).BackColor = COULEURS.BLANC
                                LNomsPostes(a).ForeColor = COULEURS.NOIR
                            End If
                        End If
                    
                    Else
                
                        '--- chariot vide verrouillé ---
                        If a = POSTES.P_CHGT_1 Or a = POSTES.P_CHGT_2 Then
                            If PBCharges(a).Picture <> PCCharges.GraphicCell(IDX_CHARIOT_PRESENT_VERROUILLE) Then
                                Set PBCharges(a).Picture = PCCharges.GraphicCell(IDX_CHARIOT_PRESENT_VERROUILLE)
                                LNomsPostes(a).BackColor = COULEURS.BLANC
                                LNomsPostes(a).ForeColor = COULEURS.NOIR
                            End If
                        Else
                            If PBCharges(a).Picture <> PCCharges.GraphicCell(IDX_CHARIOT_PRESENT_VERROUILLE + AJOUT_IMG_POSTE_CHGT_3_ET_4) Then
                                Set PBCharges(a).Picture = PCCharges.GraphicCell(IDX_CHARIOT_PRESENT_VERROUILLE + AJOUT_IMG_POSTE_CHGT_3_ET_4)
                                LNomsPostes(a).BackColor = COULEURS.BLANC
                                LNomsPostes(a).ForeColor = COULEURS.NOIR
                            End If
                        End If
                        
                    End If
                
                Case Else
                    
            End Select
                    
        End With
    
    Next a
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- contrôle de la validité du chargement ---
    If ControleValiditeChargement() = True Then
        IAutorisationChargement.Visible = True
    Else
        IAutorisationChargement.Visible = False
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Vérifie l'existence d'une commande interne dans une des tables
' Entrées :
' Retours  : ExistenceCommandeInterne ->   TRUE = La commande interne existe déjà dans la table
'                                                                     FALSE = La commande interne n'existe pas dans la table
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ExistenceCommandeInterne(ByVal GrilleDepart As TYPES_GRILLES, _
                                                                          ByVal NumCommandeInterne As Long) As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    ExistenceCommandeInterne = False
    
    If GrilleDepart = TG_CHARGEMENT Then
    
        '--- contrôle sur le chargement ---
        For a = 1 To UBound(TChargement.TDetailsCharges())
            With TChargement.TDetailsCharges(a)
                If .NumCommandeInterne = NumCommandeInterne Then
                    ExistenceCommandeInterne = True
                    Exit For
                End If
            End With
        Next a
    
    Else

        '--- contrôle sur le prévisionnel ---
        For a = 1 To NBR_LIGNES_PREVISIONNEL
            With TPrevisionnel(a)
                If .NumCommandeInterne = NumCommandeInterne Then
                    ExistenceCommandeInterne = True
                    Exit For
                End If
            End With
        Next a
    
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Contrôle la validité du chargement
' Entrées :
' Retours : ControleValiditeChargement -> TRUE = Le chargement est valide
'                                                                   FALSE = Le chargement est invalide
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ControleValiditeChargement() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    ControleValiditeChargement = True
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- contrôle pour être sûr d'avoir au moins une fiche ---
    If TChargement.TDetailsCharges(1).NumCommandeInterne = 0 Or TChargement.TDetailsCharges(1).NbrPieces = 0 Then
        ControleValiditeChargement = False
        Exit Function
    End If

    '--- contrôle sur les autres fiches ---
    For a = 2 To NBR_LIGNES_DETAILS_CHARGES
        With TChargement.TDetailsCharges(a)
            If .NumCommandeInterne = 0 Then
                Exit For
            Else
                If .NbrPieces = 0 Then
                    ControleValiditeChargement = False
                    Exit For
                End If
            End If
        End With
    Next a
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- contrôle sur le numéro de la gamme d'anodisation ---
    If TBNumGammeAnodisation.Text = "" Or GammeTrouvee = False Then
        ControleValiditeChargement = False
        Exit Function
    Else
        If IsNumeric(TBNumGammeAnodisation.Text) = True Then
        Else
            ControleValiditeChargement = False
            Exit Function
        End If
    End If

    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If FTempsEtPosteAnodisation.Visible = True Then
        
        '--- validite du champ du temps d'anodisation ---
        If MEBTempsReelAnodisation.Enabled = False Then
            ControleValiditeChargement = False
            Exit Function
        End If
        
        '--- temps d'anodisation ---
        If CTempsTexteEnSecondes(MEBTempsReelAnodisation.Text) <= 0 Then
            ControleValiditeChargement = False
            Exit Function
        End If
    
        '--- poste d'anodisation ---
        If CBChoixPosteAnodisation.ListIndex = -1 Then
            ControleValiditeChargement = False
            Exit Function
        End If

    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- contrôle si la gamme est passable en ligne ---
    If GammePassableEnLigne = False Then
        ControleValiditeChargement = False
        Exit Function
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    '--- contrôle d'un numéro de barre différent de zéro ---
    If NumBarreEnCours = 0 Then
        ControleValiditeChargement = False
        Exit Function
    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Charge la gamme d'anodisation pour le chargement et affiche les redresseurs concernés
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ChargeGammeAnodisationChargement(ByVal NumGamme As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                                                        'pour les boucles FOR...NEXT
    Dim b As Integer                                                        'pour les boucles FOR...NEXT
    Dim NumZone As Integer                                          'numéro de zone
    Dim NumPremierPoste As Integer                            'numéro du premier poste
    Dim ModeUouI As Integer                                         'pour le passage par référence
    
    '--- rendre invisible la partie redresseur et la partie du temps et poste d'anodisation ---
    GammePassableEnLigne = False                    'RAZ de la variable gamme passable en ligne
    FTempsEtPosteAnodisation.Visible = False
    FRedresseurs.Visible = False
    
    '--- vider certain champ ---
    LNomGammeAnodisation.Caption = ""
    
    '---- recherche de la gamme d'anodisation ---
    If RechercheGammesAnodisation(NumGamme) = TROUVE Then
        
        '--- affectation ---
        GammeTrouvee = True
 
        '--- transfert de la gamme d'anodisation dans le chargement ---
        TChargement.TGammesAnodisation = TTempEnrGammesAnodisation
        
        
       
        
        '--- affichage du nom de la gamme ---
        LNomGammeAnodisation.Caption = TChargement.TGammesAnodisation.NomGamme
        
        '--- affichage des options ---
        If PassageBrillantage(TTempEnrGammesAnodisation) = True Then
            SDecorationActiverAirBainBrillantage.Visible = True                                                                             'rendre visible l'activation de l'air dans le bain de brillantage
            CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).Visible = True
        Else
            SDecorationActiverAirBainBrillantage.Visible = False                                                                           'rendre invisible l'activation de l'air dans le bain de brillantage car il n'y a pas de brillantage dans la gamme
            CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).Visible = False
        End If
        FOptions.Visible = True

        '--- affichage des numéros de barres ---
        FNumBarres.Visible = True
        RepositonnerCadre.Visible = True
        
        '--- affichage des redresseurs ---
        With TChargement.TGammesAnodisation
            
            For a = LBound(.TDetailsGammesAnodisation()) To UBound(.TDetailsGammesAnodisation())
                
                '--- affectation ---
                NumZone = .TDetailsGammesAnodisation(a).NumZone
                
                If NumZone > 0 Then
                    
                    NumPremierPoste = TEtatsPostes(TZones(NumZone).NumPremierPoste).DefinitionPoste.NumPoste
                    
                    Select Case NumPremierPoste
                        
                        Case PREMIER_BAIN To POSTES.P_C12, POSTES.P_C17 To DERNIER_POSTE
                            '--- dégraissage ---
                            GammePassableEnLigne = True                    'Montée de la variable gamme passable en ligne
                        
                        Case POSTES.P_C13, POSTES.P_C14, POSTES.P_C15, POSTES.P_C16
                            '--- poste d'anodisation C13, C14, C15, C16 ---
                            GammePassableEnLigne = True                    'Montée de la variable gamme passable en ligne
                            
                            FRedresseurs.Visible = True
                            FRedresseurs.Refresh
                            
                            '--- affichage de la partie redresseur ---
                            ModeUouI = .ModeUouI
                            Call LModeUouI_Click(ModeUouI)
                    
                            '--- phases ---
                            For b = PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4
                                MEBTempsPhases(b).Text = Right(CTemps2(.TDetailsPhases(b).TempsPhase), 7)
                                TBTensionsPhases(b).Text = Format(.TDetailsPhases(b).UPhase, FORMAT_TENSION_1_DECIMALE)
                                TBIntensitesPhases(b).Text = Format(.TDetailsPhases(b).IPhase, FORMAT_INTENSITE_ENTIER)
                            Next b
                            
                            '--- rendre visible la partie du poste d'anodisation ---
                            FTempsEtPosteAnodisation.Visible = True
                            FTempsEtPosteAnodisation.Refresh

                            '--- transfert du temps de gamme et mémorisation de celui-ci ---
                            MEBTempsReelAnodisation.Text = .TDetailsGammesAnodisation(a).TempsAuPosteTexte
                            LTempsAnodisationGamme.Caption = .TDetailsGammesAnodisation(a).TempsAuPosteTexte
                            
                            '--- mémorisation des temps de l'anodisation ---
                            MemTempsReelAnodisationSecondes = .TDetailsGammesAnodisation(a).TempsAuPosteSecondes
                            MemTempsReelAnodisationTexte = .TDetailsGammesAnodisation(a).TempsAuPosteTexte
                        
                            '--- initialisation du choix du poste d'anodisation ---
                            CBChoixPosteAnodisation.ListIndex = -1
                        
                        Case Else
                    End Select
                
                Else
                    Exit For
                End If
            
            Next a
        
        End With
        
    Else
    
        '--- affectation ---
        GammeTrouvee = False
    
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des détails des charges
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsCharges(ByVal EtatSouhaite As GESTION_GRILLES)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---

    '--- déclaration ---
    Dim TypeCouleur As Boolean
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer, _
            PtrLigne As Integer
    Dim FicheVideDetailsCharges As DetailsCharges, _
            TCopieDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As DetailsCharges

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille des détails ---
            With MSHFGDetailsCharges

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_CHARGES + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_CHARGES + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 750                                        'épaisseur des titres
                .RowHeightMin = 315
                .Row = 0

                '--- paramètrages de chaque colonne ---
                .Col = COLONNES_DETAILS_CHARGES.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                .CellPictureAlignment = flexAlignCenterCenter: Set .CellPicture = ILIcones.ListImages("fleche basse").ExtractIcon

                .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                .ColWidth(.Col) = 12 * EPAISSEUR_CARACTERE: .Text = "Numéro de pointage"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                '.Col = COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                '.ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "R."
                '.ColAlignment(.Col) = flexAlignCenterCenter

                .Col = COLONNES_DETAILS_CHARGES.C_CODE_CLIENT
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Code client"
                .ColAlignment(.Col) = flexAlignLeftCenter

                .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                .ColWidth(.Col) = 7 * EPAISSEUR_CARACTERE: .Text = "Nombre de pièces"
                .ColAlignment(.Col) = flexAlignRightCenter

                .Col = COLONNES_DETAILS_CHARGES.C_DESIGNATION
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Désignation"
                .ColAlignment(.Col) = flexAlignLeftCenter

                .Col = COLONNES_DETAILS_CHARGES.C_GAMME
                .ColWidth(.Col) = 12 * EPAISSEUR_CARACTERE: .Text = "Gamme"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_MATIERE
                .ColWidth(.Col) = 12 * EPAISSEUR_CARACTERE: .Text = "Matière"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_OBSERVATIONS
                .ColWidth(.Col) = 30 * EPAISSEUR_CARACTERE: .Text = "Observations"
                .ColAlignment(.Col) = flexAlignLeftCenter

                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = LBound(TChargement.TDetailsCharges()) To UBound(TChargement.TDetailsCharges())

                    '--- N° de lignes ---
                    .Col = COLONNES_DETAILS_CHARGES.C_NUM_LIGNES
                    '.RowHeight(a) = 300                    'épaisseur des lignes
                    .Row = a
                    .Text = CStr(a)

                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_CHARGES.C_MATIERE
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.ORANGE_1, COULEURS.CYAN_1)

                    TypeCouleur = Not (TypeCouleur)

                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = LBound(TChargement.TDetailsCharges()) To UBound(TChargement.TDetailsCharges())
                TChargement.TDetailsCharges(a) = FicheVideDetailsCharges
            Next a
            With MSHFGDetailsCharges
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données dans le tableau ---

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---
            PtrLigne = 1
            For a = 1 To UBound(TChargement.TDetailsCharges())
                If TChargement.TDetailsCharges(a).NumCommandeInterne > 0 Then
                    TCopieDetailsCharges(PtrLigne) = TChargement.TDetailsCharges(a)
                    Inc PtrLigne
                End If
            Next a
            For a = 1 To UBound(TChargement.TDetailsCharges())
                TChargement.TDetailsCharges(a) = TCopieDetailsCharges(a)
            Next a

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGDetailsCharges

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False
                
                For a = LBound(TChargement.TDetailsCharges()) To UBound(TChargement.TDetailsCharges())
                    .Row = a
                    If TChargement.TDetailsCharges(a).NumCommandeInterne = 0 Then

                        TChargement.TDetailsCharges(a) = FicheVideDetailsCharges
                        For b = 1 To NBR_COLONNES_DETAILS_CHARGES
                            .Col = b
                            If .Text <> "" Then .Text = ""
                            If .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE Then
                                If .CellPicture <> LoadPicture() Then
                                    Set .CellPicture = LoadPicture()
                                End If
                            End If
                        Next b

                    Else

                        .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                        If .Text <> CStr(TChargement.TDetailsCharges(a).NumCommandeInterne) Then
                            .Text = CStr(TChargement.TDetailsCharges(a).NumCommandeInterne)

                            '--- bouton du choix des références du client ---
                            .CellPictureAlignment = flexAlignRightCenter
                            If .Text = "" Then
                                If .CellPicture <> LoadPicture() Then
                                    Set .CellPicture = LoadPicture()
                                End If
                            Else
                                If .CellPicture <> Me.ILImagesPourGrilles.ListImages(IMG_BOUTON_BAS).Picture Then
                                    Set .CellPicture = Me.ILImagesPourGrilles.ListImages(IMG_BOUTON_BAS).Picture
                                End If
                            End If

                        End If

                        '.Col = COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                        'If .Text <> TChargement.TDetailsCharges(a).NbrReparations Then .Text = TChargement.TDetailsCharges(a).NbrReparations
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_CODE_CLIENT
                        If .Text <> TChargement.TDetailsCharges(a).CodeClient Then .Text = TChargement.TDetailsCharges(a).CodeClient

                        .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                        If TChargement.TDetailsCharges(a).NbrPieces = 0 Then
                            If .Text <> "" Then .Text = ""
                        Else
                            If .Text <> CStr(TChargement.TDetailsCharges(a).NbrPieces) Then .Text = CStr(TChargement.TDetailsCharges(a).NbrPieces)
                        End If

                        .Col = COLONNES_DETAILS_CHARGES.C_DESIGNATION
                        If .Text <> TChargement.TDetailsCharges(a).Designation Then .Text = TChargement.TDetailsCharges(a).Designation
                 
                        
                      
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_GAMME
                        If .Text <> TChargement.TDetailsCharges(a).NumGamme Then .Text = TChargement.TDetailsCharges(a).NumGamme
                        
                          .Col = COLONNES_DETAILS_CHARGES.C_MATIERE
                        If .Text <> TChargement.TDetailsCharges(a).Matiere Then .Text = TChargement.TDetailsCharges(a).Matiere
                        
                          .Col = COLONNES_DETAILS_CHARGES.C_OBSERVATIONS
                        If .Text <> TChargement.TDetailsCharges(a).Observations Then .Text = TChargement.TDetailsCharges(a).Observations

                    End If

                Next a

                '--- restitution des valeurs ligne, colonne ---
                .Redraw = True
                .Row = MemLigne
                .Col = MemColonne
                .FocusRect = flexFocusHeavy

            End With

        Case Else

    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des détails des références des clients
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsReferencesClient(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NumLignePresent As Boolean
    Dim a As Integer, _
            b As Integer, _
            PtrNumLignesReferencesClient As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer
    Dim NbrTotalPieces As Long
    Dim NumCommandeInterne As Long, _
            NumLignesReferencesClient As String
    Dim TNumLignesReferencesClient As Variant
    Dim FicheVide As ImgDetailsReferencesClient
    
    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation ---
            Erase TDetailsReferencesClient()
            
            With MSHFGDetailsReferencesClient
            
                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_REFERENCES_CLIENT + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_REFERENCES_CLIENT + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 750                                        'épaisseur des titres
                .RowHeightMin = 315
                .Row = 0
                
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                .ColWidth(.Col) = 7 * EPAISSEUR_CARACTERE: .Text = "Nbr de pièces"
                .ColAlignment(.Col) = flexAlignRightCenter

                .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_REFERENCE_CLIENT
                .ColWidth(.Col) = 28.2 * EPAISSEUR_CARACTERE: .Text = "Référence du client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = LBound(TDetailsReferencesClient()) To UBound(TDetailsReferencesClient())
                
                    '--- N° de lignes ---
                    .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NUM_LIGNES
                    .Row = a
                    .Text = CStr(a)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES

                .Redraw = True

            End With
            
        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage ---

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données ---
            If MemNumLigneDetailsCharges > 0 And MemNumColonneDetailsCharges > 0 Then
                
                '--- extraction du n° de la commande interne ---
                With TChargement.TDetailsCharges(MemNumLigneDetailsCharges)
                    NumCommandeInterne = .NumCommandeInterne
                    NumLignesReferencesClient = .NumLignesReferencesClient
                End With
                
                If NumCommandeInterne > 0 Then
                
                    '--- construction du tableau des n° de lignes des références clients ---
                    'le tableau est de la forme n° de ligne puis quantité etc ...
                    If NumLignesReferencesClient <> "" Then
                        TNumLignesReferencesClient = Split(NumLignesReferencesClient, "-")
                    End If
                
                    '--- affichage des travaux ---
                    'PtrNumLignesReferencesClient = 0
                    'If RechercheTravaux(NumCommandeInterne) = TROUVE Then
                    '    For a = LBound(TTempEnrTravaux()) To UBound(TTempEnrTravaux())
                    '        With TTempEnrTravaux(a)
                    '
                    '            '--- affectation dans le tableau ---
                    '            TDetailsReferencesClient(a).NbrPieces = 0
                    '            TDetailsReferencesClient(a).ReferenceClient = .Renseignements
                    '
                    '            '--- recherche des numéros présents ---
                    '            If NumLignesReferencesClient <> "" Then
                    '                If PtrNumLignesReferencesClient <= UBound(TNumLignesReferencesClient) Then
                    '                    If a = TNumLignesReferencesClient(PtrNumLignesReferencesClient) Then
                    '                        Inc PtrNumLignesReferencesClient
                    '                        TDetailsReferencesClient(a).NbrPieces = TNumLignesReferencesClient(PtrNumLignesReferencesClient)
                    '                        Inc PtrNumLignesReferencesClient
                    '                    End If
                    '                End If
                    '            End If
                    '
                    '        End With
                    '    Next a
                    'End If
            
                End If
            
            End If
        
        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression ---

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage ---
            With MSHFGDetailsReferencesClient

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TDetailsReferencesClient()) To UBound(TDetailsReferencesClient())
                    .Row = a
                    If TDetailsReferencesClient(a).ReferenceClient = "" Then
                        
                        TDetailsReferencesClient(a) = FicheVide
                        For b = 1 To NBR_COLONNES_DETAILS_REFERENCES_CLIENT
                            .Col = b
                            If .Text <> "" Then .Text = ""
                        Next b
                    
                    Else
                        
                        .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                        If TDetailsReferencesClient(a).NbrPieces = 0 Then
                            .Text = ""
                        Else
                            .Text = CStr(TDetailsReferencesClient(a).NbrPieces)
                        End If
                        
                        .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_REFERENCE_CLIENT
                        .Text = CStr(TDetailsReferencesClient(a).ReferenceClient)
                    
                    End If
                Next a

                '--- restitution des valeurs ligne, colonne ---
                .Redraw = True
                .Row = MemLigne
                .Col = MemColonne
                .FocusRect = flexFocusHeavy

            End With

        Case GESTION_GRILLES.GG_MEMORISATION
            '--- mémorisation ---
            NumLignesReferencesClient = ""
            NbrTotalPieces = 0
            
            '--- construction de la chaine ---
            For a = LBound(TDetailsReferencesClient()) To UBound(TDetailsReferencesClient())
                If TDetailsReferencesClient(a).NbrPieces <> 0 Then
                    NbrTotalPieces = NbrTotalPieces + TDetailsReferencesClient(a).NbrPieces
                    NumLignesReferencesClient = NumLignesReferencesClient & _
                                                                      CStr(a) & "-" & _
                                                                      CStr(TDetailsReferencesClient(a).NbrPieces) & "-"
                End If
            Next a
            
            '--- enlever le dernier tiret ---
            If NumLignesReferencesClient <> "" Then
                NumLignesReferencesClient = Left(NumLignesReferencesClient, Pred(Len(NumLignesReferencesClient)))
            End If
            
            '--- mémorisation dans le chargement ---
            With TChargement.TDetailsCharges(MemNumLigneDetailsCharges)
                .NumLignesReferencesClient = NumLignesReferencesClient
                .NbrPieces = NbrTotalPieces
            End With
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet l'édition du chargement
' Entrées : KeyAscii -> Code ASCII de la touche frappée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionChargement(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- édition uniquement sur les bonnes colonnes ---
    Select Case MSHFGDetailsCharges.Col

        Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE, _
                                 COLONNES_DETAILS_CHARGES.C_NBR_PIECES
             ' COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS,
            With MEBEditionDetailsCharges

                '--- affiche le contrôle texte au bon endroit (dans la cellule) ---
                .Move MSHFGDetailsCharges.Left + MSHFGDetailsCharges.CellLeft, _
                           MSHFGDetailsCharges.Top + MSHFGDetailsCharges.CellTop, _
                           MSHFGDetailsCharges.CellWidth, _
                           MSHFGDetailsCharges.CellHeight

                '--- paramètres de contrôle texte en fonction de la cellule ---
                .Mask = ""
                .Text = ""
                Select Case MSHFGDetailsCharges.Col
                    'SZP 2023
                    Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE: .Mask = ""
                    'Case COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS: .Mask = "#"
                    Case COLONNES_DETAILS_CHARGES.C_NBR_PIECES: .Mask = "######"
                    Case Else
                End Select

                '--- analyse du caractère qui a été tapé ---
                Select Case KeyAscii

                    Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                        '--- du code 0 à l'espace (sauf retour arrière, Entrée) cela signifie une modification du texte en cours ---
                        .SelText = MSHFGDetailsCharges.Text
                        .SelStart = 0
                        .SelLength = Len(Replace(.Text, "_", ""))
                        .Visible = True
                        .SetFocus

                    Case vbKeyBack
                        '--- touche retour arrière ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        DoEvents
                        MEBEditionDetailsCharges_Change
                    
                    Case vbKeyReturn
                        '--- touche Entrée ---
                        With MSHFGDetailsCharges
                            Select Case .Col
                                Case COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE: .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                                'Case COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS: .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                                Case COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                                    If .Row < .Rows - 1 Then .Row = .Row + 1
                                    .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                                Case Else
                            End Select
                        End With

                    Case Else
                        '--- tout autre élément signifie le remplacement du texte en cours ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        SendKeys Chr(KeyAscii)

                End Select

            End With

        Case Else

    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet l'édition du prévisionnel
' Entrées : KeyAscii -> Code ASCII de la touche frappée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionPrevisionnel(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- édition uniquement sur les bonnes colonnes ---
    Select Case MSHFGPrevisionnel.Col

        Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE, _
                 COLONNES_PREVISIONNEL.C_NBR_REPARATIONS, _
                 COLONNES_PREVISIONNEL.C_NBR_PIECES, _
                 COLONNES_PREVISIONNEL.C_NUM_BARRE, _
                 COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION, _
                 COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION

            With MEBEditionPrevisionnel

                '--- affiche le contrôle texte au bon endroit (dans la cellule) ---
                .Move MSHFGPrevisionnel.Left + MSHFGPrevisionnel.CellLeft, _
                           MSHFGPrevisionnel.Top + MSHFGPrevisionnel.CellTop, _
                           MSHFGPrevisionnel.CellWidth, _
                           MSHFGPrevisionnel.CellHeight

                '--- paramètres de contrôle texte en fonction de la cellule ---
                .Mask = ""
                .Text = ""
                Select Case MSHFGPrevisionnel.Col
                    Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE: .Mask = "########"
                    Case COLONNES_PREVISIONNEL.C_NBR_REPARATIONS: .Mask = "#"
                    Case COLONNES_PREVISIONNEL.C_NBR_PIECES: .Mask = "######.##"
                    Case COLONNES_PREVISIONNEL.C_NUM_BARRE: .Mask = "##"
                    Case COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION: .Mask = "######"
                    Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION: .Mask = "#"
                    Case Else
                End Select

                '--- analyse du caractère qui a été tapé ---
                Select Case KeyAscii

                    Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                        '--- du code 0 à l'espace (sauf retour arrière, Entrée) cela signifie une modification du texte en cours ---
                        .SelText = MSHFGPrevisionnel.Text
                        .SelStart = 0
                        .SelLength = Len(Replace(.Text, "_", ""))
                        .Visible = True
                        .SetFocus

                    Case vbKeyBack
                        '--- touche retour arrière ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        DoEvents
                        MEBEditionPrevisionnel_Change
                    
                    Case vbKeyReturn
                        '--- touche Entrée ---
                        With MSHFGPrevisionnel
                            Select Case .Col
                                Case COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE: .Col = COLONNES_PREVISIONNEL.C_NBR_PIECES
                                Case COLONNES_PREVISIONNEL.C_NBR_REPARATIONS: .Col = COLONNES_PREVISIONNEL.C_NBR_PIECES
                                Case COLONNES_PREVISIONNEL.C_NBR_PIECES: .Col = COLONNES_PREVISIONNEL.C_NUM_BARRE
                                Case COLONNES_PREVISIONNEL.C_NUM_BARRE: .Col = COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                                Case COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION: .Col = COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                                Case COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                                    If .Row < .Rows - 1 Then .Row = .Row + 1
                                    .Col = COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                                Case Else
                            End Select
                        End With

                    Case Else
                        '--- tout autre élément signifie le remplacement du texte en cours ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        SendKeys Chr(KeyAscii)

                End Select

            End With

        Case Else

    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet l'édition du prévisionnel
' Entrées : KeyAscii -> Code ASCII de la touche frappée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionDetailsReferencesClient(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- édition uniquement sur les bonnes colonnes ---
    Select Case MSHFGDetailsReferencesClient.Col

        Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES

            With MEBEditionDetailsReferencesClient

                '--- affiche le contrôle texte au bon endroit (dans la cellule) ---
                .Move MSHFGDetailsReferencesClient.Left + MSHFGDetailsReferencesClient.CellLeft, _
                           MSHFGDetailsReferencesClient.Top + MSHFGDetailsReferencesClient.CellTop, _
                           MSHFGDetailsReferencesClient.CellWidth, _
                           MSHFGDetailsReferencesClient.CellHeight

                '--- paramètres de contrôle texte en fonction de la cellule ---
                .Mask = ""
                .Text = ""
                Select Case MSHFGDetailsReferencesClient.Col
                    Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES: .Mask = "#####.#"
                    Case Else
                End Select
                
                '--- analyse du caractère qui a été tapé ---
                Select Case KeyAscii

                    Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                        '--- du code 0 à l'espace (sauf retour arrière, Entrée) cela signifie une modification du texte en cours ---
                        .SelText = MSHFGDetailsReferencesClient.Text
                        .SelStart = 0
                        .SelLength = Len(Replace(.Text, "_", ""))
                        .Visible = True
                        .SetFocus

                    Case vbKeyBack
                        '--- touche retour arrière ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        DoEvents
                        MEBEditionDetailsReferencesClient_Change

                    Case vbKeyReturn
                        '--- touche Entrée ---
                        With MSHFGDetailsReferencesClient
                            Select Case .Col
                                Case COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                                    If .Row < .Rows - 1 Then .Row = .Row + 1
                                    .Col = COLONNES_DETAILS_REFERENCES_CLIENT.C_NBR_PIECES
                                Case Else
                            End Select
                        End With

                    Case Else
                        '--- tout autre élément signifie le remplacement du texte en cours ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        SendKeys Chr(KeyAscii)

                End Select

            End With
        
        Case Else

    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Insertion d'une commande interne dans une des grilles
' Entrées :             GrilleConcernee -> Fonction de l'énumération TYPES_GRILLES
'                                      NumLigne -> N° de ligne dans le tableau
'                 NumCommandeInterne -> N° de la commande interne
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function InsertionCommandeInterne(ByVal GrilleConcernee As TYPES_GRILLES, _
                                                                        ByVal NumLigne As Integer, _
                                                                        ByVal NumCommandeInterne As Long) As String

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- lancer la modification ---
    GestionBoutons E_MODIFICATION_EN_COURS

    
    
    '--- Base de données CLIPPER ---
    If NumCommandeInterne > 0 Then
    
    
        TBNumGammeAnodisation = TTempEnrPhasesClipper.NumGamme
        TBMatiere = TTempEnrPhasesClipper.Matiere
        
        
        Dim NumGammeTexte As String
        NumGammeTexte = Format(CLng(TBNumGammeAnodisation.Text), FORMAT_NUM_GAMME_ANODISATION)
            
        If RecherchePhasesClipper(NumCommandeInterne) = TROUVE Then
            
            
            ChargeGammeAnodisationChargement NumGammeTexte
            
            If GrilleConcernee = TG_CHARGEMENT Then
            
                '--- grille du chargement ---
                With TChargement.TDetailsCharges(NumLigne)
                    .NumCommandeInterne = TTempEnrPhasesClipper.GaCLeUnik
                    .NumGamme = TTempEnrPhasesClipper.NumGamme
                    .Naf = TTempEnrPhasesClipper.Naf
                    .CodeClient = TTempEnrPhasesClipper.CoCli
                    .Designation = TTempEnrPhasesClipper.Desa1
                    .Observations = TTempEnrPhasesClipper.GamObs
                    .Matiere = TTempEnrPhasesClipper.Matiere
                    .NumLignesReferencesClient = ""
                End With
        
            Else
        
                '--- grille du prévisionnel ---
                With TPrevisionnel(NumLigne)
                    .NumCommandeInterne = TTempEnrPhasesClipper.GaCLeUnik
                    .CodeClient = TTempEnrPhasesClipper.CoCli
                    .NumGamme = TTempEnrPhasesClipper.NumGamme
                    .NbrPieces = TTempEnrPhasesClipper.QteAf
                    .Designation = TTempEnrPhasesClipper.Desa1
                    .Observations = TTempEnrPhasesClipper.GamObs
                    .Matiere = TTempEnrPhasesClipper.Matiere
                End With
        
            End If
            
            '--- affectation ---
            InsertionCommandeInterne = TROUVE
            'MsgBox ("Clipper trouvé")
            
      Else

           '--- affectation -
           '-- SZP 20241004 TODO CHECK
           InsertionCommandeInterne = NON_TROUVE
           
           
           If GrilleConcernee = TG_CHARGEMENT Then
            
                '--- grille du chargement ---
                With TChargement.TDetailsCharges(NumLigne)
                    .NumCommandeInterne = NumCommandeInterne
                    .NumGamme = NumGammeTexte
                    .Matiere = TBMatiere
                    .NumLignesReferencesClient = ""
                End With
        
            Else
        
                '--- grille du prévisionnel ---
                With TPrevisionnel(NumLigne)
                    .NumCommandeInterne = NumCommandeInterne
                    
                    .NumGamme = NumGammeTexte
                    .Matiere = TBMatiere
                End With
        
            End If
   
      End If
            
    
    

    
    Else
            
        '--- affectation ---
        InsertionCommandeInterne = NON_TROUVE
    
    End If

   
    
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Insère une ligne sur l'une des grilles
' Entrées : GrilleConcernee -> Fonction de l'énumération TYPES_GRILLES
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InsererLigneGrille(ByVal GrilleConcernee As TYPES_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer, _
           NumLigne As Integer
    Dim FicheVideDetailsCharges As DetailsCharges, _
            FicheVidePrevisionnel As VarPrevisionnel
    
    If GrilleConcernee = TG_CHARGEMENT Then
    
        '--- affectation ---
        NumLigne = MSHFGDetailsCharges.Row

        '--- insertion de la ligne ---
        If NumLigne > 0 And NumLigne <= NBR_LIGNES_DETAILS_CHARGES Then
            For a = Pred(NBR_LIGNES_DETAILS_CHARGES) To NumLigne Step -1
                TChargement.TDetailsCharges(Succ(a)) = TChargement.TDetailsCharges(a)
            Next a
            TChargement.TDetailsCharges(NumLigne) = FicheVideDetailsCharges
            GestionDetailsCharges GG_AFFICHAGE
            With MSHFGDetailsCharges
                .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                .SetFocus
            End With
        End If

    Else

        '--- affectation ---
        NumLigne = MSHFGPrevisionnel.Row

        '--- insertion de la ligne ---
        If NumLigne > 0 And NumLigne <= NBR_LIGNES_PREVISIONNEL Then
            For a = Pred(NBR_LIGNES_PREVISIONNEL) To NumLigne Step -1
                TPrevisionnel(Succ(a)) = TPrevisionnel(a)
            Next a
            TPrevisionnel(NumLigne) = FicheVidePrevisionnel
            GestionPrevisionnel GG_AFFICHAGE
            With MSHFGPrevisionnel
                .Col = COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                .SetFocus
            End With
        End If

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Compacte une des grilles
' Entrées : GrilleConcernee -> Fonction de l'énumération TYPES_GRILLES
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CompacterGrille(ByVal GrilleConcernee As TYPES_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- compactage ---
    If GrilleConcernee = TG_CHARGEMENT Then
        GestionDetailsCharges GG_COMPRESSION
        GestionDetailsCharges GG_AFFICHAGE
    Else
        GestionPrevisionnel GG_COMPRESSION
        GestionPrevisionnel GG_AFFICHAGE
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Transfert les données d'une grille vers l'autre grille
' Entrées : GrilleDepart         -> Fonction de l'énumération TYPES_GRILLES
'                 NumLigneDepart -> N° de ligne faisant l'objet du transfert
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub TransfertDonneesEntreGrilles(ByVal GrilleDepart As TYPES_GRILLES, _
                                                                     ByVal NumLigneDepart As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    If GrilleDepart = TG_CHARGEMENT Then
    
    Else
    
        '--- transfert des données vers le chargement ---
        If TChargement.TDetailsCharges(NBR_LIGNES_DETAILS_CHARGES).NumCommandeInterne = 0 And _
            NumLigneDepart > 0 Then
            
            '--- transfert ---
            With TChargement.TDetailsCharges(NBR_LIGNES_DETAILS_CHARGES)
                .NumCommandeInterne = TPrevisionnel(NumLigneDepart).NumCommandeInterne
                .NbrReparations = TPrevisionnel(NumLigneDepart).NbrReparations
                .CodeClient = TPrevisionnel(NumLigneDepart).CodeClient
                .NbrPieces = TPrevisionnel(NumLigneDepart).NbrPieces
                .Designation = TPrevisionnel(NumLigneDepart).Designation
                .Matiere = TPrevisionnel(NumLigneDepart).Matiere
            End With
                    
            '--- rafraichissement ---
            GestionDetailsCharges GG_COMPRESSION
            GestionDetailsCharges GG_AFFICHAGE
                    
        End If
                    
    End If
                    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Supprime une ligne sur l'une des grilles
' Entrées : GrilleConcernee -> Fonction de l'énumération TYPES_GRILLES
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SupprimerLigneGrille(ByVal GrilleConcernee As TYPES_GRILLES)
  
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumLigne As Integer
    Dim FicheVideDetailsCharges As DetailsCharges, _
            FicheVidePrevisionnel As VarPrevisionnel
    
    If GrilleConcernee = TG_CHARGEMENT Then
    
        '--- affectation ---
        NumLigne = MSHFGDetailsCharges.Row

        '--- suppression de la ligne ---
        If NumLigne > 0 And NumLigne <= NBR_LIGNES_DETAILS_CHARGES Then
            If AppelFenetre(F_MESSAGE, _
                                    TITRE_MESSAGES, _
                                    vbCrLf & "cs|Suppression d'une ligne dans le CHARGEMENT" & _
                                    MESSAGE_3 & CStr(NumLigne) & " ?", _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                TChargement.TDetailsCharges(NumLigne) = FicheVideDetailsCharges
                GestionDetailsCharges GG_COMPRESSION
                GestionDetailsCharges GG_AFFICHAGE
            End If
            MSHFGDetailsCharges.SetFocus
        End If

    Else

        '--- affectation ---
        NumLigne = MSHFGPrevisionnel.Row

        '--- suppression de la ligne ---
        If NumLigne > 0 And NumLigne <= NBR_LIGNES_PREVISIONNEL Then
            If AppelFenetre(F_MESSAGE, _
                                    TITRE_MESSAGES, _
                                    vbCrLf & "cs|Suppression d'une ligne dans le PREVISIONNEL" & _
                                    MESSAGE_3 & CStr(NumLigne) & " ?", _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                TPrevisionnel(NumLigne) = FicheVidePrevisionnel
                GestionPrevisionnel GG_COMPRESSION
                GestionPrevisionnel GG_AFFICHAGE
            End If
            MSHFGPrevisionnel.SetFocus
        End If

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Introduit une charge dans la ligne (transfert des valeurs de chargement dans l'automate)
' Entrées : NumPoste -> Numéro du poste ou se trouve la charge
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub IntroductionChargeAuChargement(ByVal NumPoste As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
                
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer, _
           NumCharge As Integer, _
           NumZone As Integer, _
           NumPremierPoste As Integer
    Dim ValeurRetourneeAPI As Long                  'valeur retournée par une fonction concernant le dialogue avec l'automate
    
    '--- introduction de la charge ---
   
    If NumPoste >= POSTES.P_CHGT_1 Then 'And NumPoste <= POSTES.P_CHGT_2 Then
        
        If TEtatsPostes(NumPoste).NumCharge = 0 Then
                    
            '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- recherche du prochain numéro de charge ---
            NumCharge = ProchainNumeroCharge()
    
            '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- initialisation pour être sûr de vider toutes les données ---
            InitialisationCharge NumCharge
    
            '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            '--- remplir la totalité de la fiche ---
            With TEtatsCharges(NumCharge)
                
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- date d'entrée en ligne ---
                .DateEntreeEnLigne = Now
                
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- numéro de barre ---
                .NumBarre = NumBarreEnCours
                'SZP 2021
                .NumBarreInc = getIDBARRE()
               
                
                
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- transfert du temps en secondes du délai de stabilisation de la charge ---
                If IsNumeric(TBDelaiSupStabilisationCharge.Text) = True Then
                    .DelaiSupStabilisationChargeSecondes = CInt(TBDelaiSupStabilisationCharge.Text)
                Else
                    .DelaiSupStabilisationChargeSecondes = 0
                End If
                
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                'Call LogCharges("************************************")
                'Call LogCharges("numéro de barre : " & .NumBarreInc)
                '--- transfert des détails de la charge ---
                For a = 1 To UBound(.TDetailsCharges())
                    .TDetailsCharges(a) = TChargement.TDetailsCharges(a)
                    
                    If .TDetailsCharges(a).NumCommandeInterne > 0 Then
                    
                      'Call LogCharges("Chargement NAF: " & .TDetailsCharges(a).Naf & " , nb pieces  : " & .TDetailsCharges(a).NbrPieces & _
                      '  " , client  : " & .TDetailsCharges(a).CodeClient & ", pointage:" & .TDetailsCharges(a).NumCommandeInterne)
                    End If
                        
                Next a
                
                'Call LogCharges("************************************")
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--- transfert de la gamme d'anodisation ---
                .TGammesAnodisation = TChargement.TGammesAnodisation
                
                'TMP TEST
                '.DateArriveeAuDechargement = DateAdd("s", 4607, Now)
                
                
                'insertionClipperPointage (NumCharge)
                
                
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                Dim NumDerniereLigneGamme As Integer
                
                '--- transfert du temps réel d'anodisation ---
                For a = LBound(.TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(.TGammesAnodisation.TDetailsGammesAnodisation())
                    With .TGammesAnodisation.TDetailsGammesAnodisation(a)
                        NumZone = .NumZone
                        If NumZone > 0 Then
                            NumDerniereLigneGamme = a
                            NumPremierPoste = TEtatsPostes(TZones(NumZone).NumPremierPoste).DefinitionPoste.NumPoste
                            Select Case NumPremierPoste
                                Case POSTES.P_C13
                                    If MEBTempsReelAnodisation.Text <> .TempsAuPosteTexte Then
                                        .TempsAuPosteTexte = MEBTempsReelAnodisation.Text
                                        .TempsAuPosteSecondes = CTempsTexteEnSecondes(.TempsAuPosteTexte)
                                    End If
                                
                                Case Else
                            End Select
                        Else
                            Exit For
                        End If
                    End With
                Next a
                
                
                 '--- choix du poste d'anodisation ---
                If CBChoixPosteAnodisation.ListIndex = -1 Then
                    .TGammesAnodisation.ChoixPosteAnodisation = C_AUTOMATIQUE
                Else
                    .TGammesAnodisation.ChoixPosteAnodisation = CBChoixPosteAnodisation.ListIndex
                End If
                
                
                If CBOptionsEtuve(1).value = 1 Then
                    'on décale la zone de déchargement une ligne plus haut
                    .TGammesAnodisation.TDetailsGammesAnodisation(NumDerniereLigneGamme + 1) = .TGammesAnodisation.TDetailsGammesAnodisation(NumDerniereLigneGamme)
                    
                    ' l'avant dernière zone devient un passage à l'étuve
                    With .TGammesAnodisation.TDetailsGammesAnodisation(NumDerniereLigneGamme)
                        .NumZone = ZONE_ETUVE
                        .TempsAuPosteTexte = "00:" & EtuveTpsPoste.Text & ":00"
                        .TempsAuPosteSecondes = CInt(EtuveTpsPoste.Text) * 60
                        
                        
                    End With
                    
                  
                End If
                
                
        
      
                
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                If FRedresseurs.Visible = True Then
                
                    '--- mode U ou I ---
                    .ModeUouI = ModeUouIEnCours
                    
                    '--- temps des phases, tension et intensité ---
                    For a = LBound(.TDetailsPhasesProduction()) To UBound(.TDetailsPhasesProduction())
                    
                        '--- temps de la phase ---
                        .TDetailsPhasesProduction(a).TempsPhase = CTempsTexteEnSecondes(MEBTempsPhases(a).Text)
                    
                        '--- tension ---
                        If IsNumeric(TBTensionsPhases(a).Text) = True Then
                            .TDetailsPhasesProduction(a).UPhase = CInt(CSng(TBTensionsPhases(a).Text) * 10)
                        Else
                            .TDetailsPhasesProduction(a).UPhase = 0
                        End If
                        
                        '--- intensité ---
                        If IsNumeric(TBIntensitesPhases(a).Text) = True Then
                            .TDetailsPhasesProduction(a).IPhase = CInt(CSng(TBIntensitesPhases(a).Text))
                        Else
                            .TDetailsPhasesProduction(a).IPhase = 0
                        End If
                    
                    Next a
                    
                    '--- temps total de la gamme redresseur en secondes ---
                    .TempsTotalGammeRedresseur = CalculTempsTotalGammeRedresseur()
                    
                Else
                    
                    '--- fixer les valeurs par défaut ---
                    .ModeUouI = M_TENSION
                    Erase .TDetailsPhasesProduction()
                    .TempsTotalGammeRedresseur = 0
                
                End If
                    
                '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--- construction du mot des options 1 (mot transmis à l'automate) ---
                '                           Poids FORT du mot transmis
                '                           ---------------------------------------------------------------------------------------
                '                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
                '                           ---------------------------------------------------------------------------------------
                '                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
                '                           ---------------------------------------------------------------------------------------
                '                           |           |          |         |         |          |          |         |_____  forcer la montée en très petite vitesse
                '                           |           |          |         |         |          |          |__________  forcer la montée en petite vitesse
                '                           |           |          |         |         |          |________________ forcer la descente en très petite vitesse
                '                           |           |          |         |         |_____________________  forcer la descente en petite vitesse
                '                           |           |          |         |__________________________
                '                           |           |          |_______________________________
                '                           |           |_____________________________________
                '                           |___________________________________________
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                '--- construction du mot des options 2 (mot transmis à l'automate) ---
                '                           Poids FORT du mot transmis
                '                           ---------------------------------------------------------------------------------------
                '                           |  Bit 7 |  Bit 6 | Bit 5 | Bit 4 | Bit 3 | Bit 2 | Bit 1 | Bit 0 |
                '                           ---------------------------------------------------------------------------------------
                '                           |  128   |   64   |   32  |   16   |    8   |    4    |    2   |     1   |
                '                           ---------------------------------------------------------------------------------------
                '                           |           |          |         |         |          |          |         |_____  activer l'air dans le brillantage
                '                           |           |          |         |         |          |          |__________
                '                           |           |          |         |         |          |________________
                '                           |           |          |         |         |_____________________
                '                           |           |          |         |__________________________
                '                           |           |          |_______________________________
                '                           |           |_____________________________________
                '                           |___________________________________________
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                '--- initialisation du mot contenant les options 1 et 2 ---
                .Options1 = 0
                .Options2 = 0
                
                '--- options 1 / partie concernant les ponts ---
                If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_PV).value = 1 Then
                    .Options1 = .Options1 + 8                         'bit 3 du mot des options 1
                End If
                If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_DESCENTE_EN_TPV).value = 1 Then
                    .Options1 = .Options1 + 4                         'bit 2 du mot des options 1
                End If
                If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_PV).value = 1 Then
                    .Options1 = .Options1 + 2                         'bit 1 du mot des options 1
                End If
                If CBOptionsPonts(OPTIONS_GAMME.O_FORCER_MONTEE_EN_TPV).value = 1 Then
                    .Options1 = .Options1 + 1                         'bit 0 du mot des options 1
                End If
            
                '--- options 2 / partie concernant les postes ---
                If CBOptionsPostes(OPTIONS_GAMME.O_ACTIVER_AIR_BRILLANTAGE).value = 1 Then
                    .Options2 = .Options2 + 1                         'bit 0 du mot des options 2
                End If
                
            End With
                
            'DEBUT SZP 20180605
            'If TypeBD = TYPES_BD.BD_CLIPPER Then
            '   EnregistrementBainsPourCLIPPER NumCharge
            'End If
            'EnregistrementProduction (NumCharge)
            ' FIN SZP 20180605
                
            '--- envoi dans l'automate du numéro de charge ---
            If PROGRAMME_AVEC_AUTOMATE = True Then
            
                
                '--- transfert dans l'automate ---
                If EnvoiNumeroChargePosteAvecOptions(NumPoste, NumCharge, TEtatsCharges(NumCharge).Options1, TEtatsCharges(NumCharge).Options2) = OK Then
                                       
                    
                    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                    '--- lancer l'affichage des données transmises (compteur) ---
                    LDonneesTransmisesAPI.Visible = True
                    
                    '--- transfert de la charge dans l'automate ---
                    ValeurRetourneeAPI = APITransfertCharge(NumCharge, TEtatsCharges(NumCharge))
                    
                    '--- lancer un message d'erreur ---
                    If ValeurRetourneeAPI <> 0 Then
                        Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                    End If

                    '--- rendre invisible l'affichage des données transmises (compteur) ---
                    LDonneesTransmisesAPI.Visible = False
                    
                    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                    If EntreeAutomatiqueCharges = False Then
                        
                        TEtatsCharges(NumCharge).PtrZoneGammeAnodisation = 1
                        'pointer la première zone pour indiquer un début de traitement
                        '1 pointe normallement le chargement
                        'uniquement en mode d'entrées des charges en manuel,
                        'si le mode d'entrées des charges est sur automatique,
                        'c 'est le moteur d'inférence qui force le pointeur à 1
                    
                    End If
                    
                    '---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                    '--- effacement complet du chargement ---
                    EffacementCompletChargement
                
                Else
                    
                    Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_500)
                
                End If
                
            End If
            
        End If
            
    End If
    
    

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion du prévisionnel
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionPrevisionnel(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim TypeCouleur As Boolean
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer, _
            PtrLigne As Integer
    Dim UnTexte As String
    Dim FicheVide As VarPrevisionnel, _
            TCopiePrevisionnel(1 To NBR_LIGNES_PREVISIONNEL) As VarPrevisionnel

    Select Case EtatSouhaite

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille des détails ---
            With MSHFGPrevisionnel

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_PREVISIONNEL + .FixedRows
                .Cols = NBR_COLONNES_PREVISIONNEL + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 750                                        'épaisseur des titres
                .RowHeightMin = 315
                .Row = 0
                
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_PREVISIONNEL.C_NUM_LIGNES
                .ColWidth(.Col) = 5 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_PREVISIONNEL.C_CHOIX_IA
                .ColWidth(.Col) = 5.7 * EPAISSEUR_CARACTERE: .Text = "Choix"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_PREVISIONNEL.C_NBR_REPARATIONS
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "R."
                .ColAlignment(.Col) = flexAlignCenterCenter

                .Col = COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "N° de pointage"
                .ColAlignment(.Col) = flexAlignCenterCenter

                .Col = COLONNES_PREVISIONNEL.C_CODE_CLIENT
                .ColWidth(.Col) = 12.6 * EPAISSEUR_CARACTERE: .Text = "Code client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_PREVISIONNEL.C_NBR_PIECES
                .ColWidth(.Col) = 8 * EPAISSEUR_CARACTERE: .Text = "Nombre de pièces"
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_PREVISIONNEL.C_DESIGNATION
                .ColWidth(.Col) = 27 * EPAISSEUR_CARACTERE: .Text = "Désignation"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_PREVISIONNEL.C_NUM_BARRE
                .ColWidth(.Col) = 6 * EPAISSEUR_CARACTERE: .Text = "N° de barre"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                .ColWidth(.Col) = 8 * EPAISSEUR_CARACTERE: .Text = "N° de gamme"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_PREVISIONNEL.C_PASSAGE_ANODISATION
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "A."
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_PREVISIONNEL.C_PASSAGE_SPECTRO
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "S."
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_PREVISIONNEL.C_PASSAGE_OR
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "O."
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_PREVISIONNEL.C_PASSAGE_NOIR
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "N."
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                .ColWidth(.Col) = 20.3 * EPAISSEUR_CARACTERE: .Text = "Choix du poste d'anodisation"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = LBound(TPrevisionnel()) To UBound(TPrevisionnel())
                
                    '--- N° de lignes ---
                    .Col = COLONNES_PREVISIONNEL.C_NUM_LIGNES
                    '.RowHeight(a) = 300                    'épaisseur des lignes
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_PREVISIONNEL.C_CHOIX_IA
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.ORANGE_1, COULEURS.CYAN_1)
                    
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE

                .Redraw = True

            End With

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = LBound(TPrevisionnel()) To UBound(TPrevisionnel())
                TPrevisionnel(a) = FicheVide
            Next a
            With MSHFGPrevisionnel
                .TopRow = 1
                .LeftCol = 1
            End With

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données dans le tableau ---

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---
            PtrLigne = 1
            For a = 1 To NBR_LIGNES_PREVISIONNEL
                If TPrevisionnel(a).NumCommandeInterne > 0 Then
                    TCopiePrevisionnel(PtrLigne) = TPrevisionnel(a)
                    Inc PtrLigne
                End If
            Next a
            For a = 1 To NBR_LIGNES_PREVISIONNEL
                TPrevisionnel(a) = TCopiePrevisionnel(a)
            Next a

        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        
        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGPrevisionnel

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TPrevisionnel()) To UBound(TPrevisionnel())
                    
                    .Row = a
                    
                    If TPrevisionnel(a).NumCommandeInterne = 0 Then
                        
                        TPrevisionnel(a) = FicheVide
                        For b = 1 To NBR_COLONNES_PREVISIONNEL
                            .Col = b
                            If .Text <> "" Then .Text = ""
                            Select Case b
                                Case COLONNES_PREVISIONNEL.C_CHOIX_IA, COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION, COLONNES_PREVISIONNEL.C_PASSAGE_ANODISATION To COLONNES_PREVISIONNEL.C_PASSAGE_NOIR
                                    If .CellPicture <> LoadPicture() Then
                                        Set .CellPicture = LoadPicture()
                                    End If
                                Case Else
                            End Select
                        Next b
                    
                    Else
                        
                        .Col = COLONNES_PREVISIONNEL.C_CHOIX_IA
                        If TPrevisionnel(a).ChoixIA = 0 Then
                            
                            '--- effacement de l'image du choix ---
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        
                        Else
                           
                            '--- affichage de l'image du numéro ---
                            .CellPictureAlignment = flexAlignCenterCenter
                            If .CellPicture <> Me.ILImagesNumChoix.ListImages("choix " & TPrevisionnel(a).ChoixIA).Picture Then
                                Set .CellPicture = Me.ILImagesNumChoix.ListImages("choix " & TPrevisionnel(a).ChoixIA).Picture
                            End If
                            
                        End If
                        
                        .Col = COLONNES_PREVISIONNEL.C_NUM_COMMANDE_INTERNE
                        AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).NumCommandeInterne
                        
                        .Col = COLONNES_PREVISIONNEL.C_NBR_REPARATIONS
                        AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).NbrReparations
                        
                        .Col = COLONNES_PREVISIONNEL.C_CODE_CLIENT
                        AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).CodeClient
                        
                        .Col = COLONNES_PREVISIONNEL.C_NBR_PIECES
                        If TPrevisionnel(a).NbrPieces = 0 Then
                            AffichageTexte MSHFGPrevisionnel, ""
                        Else
                            AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).NbrPieces
                        End If

                        .Col = COLONNES_PREVISIONNEL.C_DESIGNATION
                        AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).Designation
                        
                        .Col = COLONNES_PREVISIONNEL.C_NUM_BARRE
                        If TPrevisionnel(a).NumBarre = 0 Then
                            AffichageTexte MSHFGPrevisionnel, ""
                        Else
                            AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).NumBarre
                        End If
                        
                        .Col = COLONNES_PREVISIONNEL.C_NUM_GAMME_ANODISATION
                        If TPrevisionnel(a).NumGammeAnodisation = "" Then
                            AffichageTexte MSHFGPrevisionnel, ""
                        Else
                            AffichageTexte MSHFGPrevisionnel, TPrevisionnel(a).NumGammeAnodisation
                        End If
                        
                        '--- indicateur de passage en anodisation ---
                        .Col = COLONNES_PREVISIONNEL.C_PASSAGE_ANODISATION
                        .CellPictureAlignment = flexAlignCenterCenter
                        If TPrevisionnel(a).TGammesAnodisation.PassageAnodisation = True Then
                            If .CellPicture <> Me.ILImagesColorations.ListImages(IMG_ANODISATION).Picture Then
                                Set .CellPicture = Me.ILImagesColorations.ListImages(IMG_ANODISATION).Picture
                            End If
                        Else
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        End If
                        
                        '--- indicateur de passage en spectrocoloration ---
                        .Col = COLONNES_PREVISIONNEL.C_PASSAGE_SPECTRO
                        .CellPictureAlignment = flexAlignCenterCenter
                        If TPrevisionnel(a).TGammesAnodisation.PassageSpectro = True Then
                            If .CellPicture <> Me.ILImagesColorations.ListImages(IMG_SPECTRO).Picture Then
                                Set .CellPicture = Me.ILImagesColorations.ListImages(IMG_SPECTRO).Picture
                            End If
                        Else
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        End If
                        
                        '--- indicateur de passage dans le bain d'or ---
                        .Col = COLONNES_PREVISIONNEL.C_PASSAGE_OR
                        .CellPictureAlignment = flexAlignCenterCenter
                        If TPrevisionnel(a).TGammesAnodisation.PassageOr = True Then
                            If .CellPicture <> Me.ILImagesColorations.ListImages(IMG_OR).Picture Then
                                Set .CellPicture = Me.ILImagesColorations.ListImages(IMG_OR).Picture
                            End If
                        Else
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        End If
                        
                        '--- indicateur de passage dans le bain dz noir ---
                        .Col = COLONNES_PREVISIONNEL.C_PASSAGE_NOIR
                        .CellPictureAlignment = flexAlignCenterCenter
                        If TPrevisionnel(a).TGammesAnodisation.PassageNoir = True Then
                            If .CellPicture <> Me.ILImagesColorations.ListImages(IMG_NOIR).Picture Then
                                Set .CellPicture = Me.ILImagesColorations.ListImages(IMG_NOIR).Picture
                            End If
                        Else
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        End If
                        
                        '--- bouton du choix du poste d'anodisation ---
                        .Col = COLONNES_PREVISIONNEL.C_CHOIX_POSTE_ANODISATION
                        .CellPictureAlignment = flexAlignRightCenter
                        If TPrevisionnel(a).NumGammeAnodisation = "" Then
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        Else
                            If .CellPicture <> Me.ILImagesPourGrilles.ListImages(IMG_BOUTON_BAS).Picture Then
                                Set .CellPicture = Me.ILImagesPourGrilles.ListImages(IMG_BOUTON_BAS).Picture
                            End If
                        End If
                        
                        '--- texte du choix du poste d'anodisation ---
                        If TPrevisionnel(a).NumGammeAnodisation = "" Then
                            AffichageTexte MSHFGPrevisionnel, ""
                        Else
                            UnTexte = Switch(TPrevisionnel(a).ChoixPosteAnodisation = CHOIX_POSTE_ANODISATION.C_AUTOMATIQUE, "AUTOMATIQUE", _
                                                         TPrevisionnel(a).ChoixPosteAnodisation = CHOIX_POSTE_ANODISATION.C_C13_IMPOSE, "C13 IMPOSE", _
                                                         TPrevisionnel(a).ChoixPosteAnodisation = CHOIX_POSTE_ANODISATION.C_C14_IMPOSE, "C14 IMPOSE", _
                                                         TPrevisionnel(a).ChoixPosteAnodisation = CHOIX_POSTE_ANODISATION.C_C15_IMPOSE, "C15 IMPOSE", _
                                                         TPrevisionnel(a).ChoixPosteAnodisation = CHOIX_POSTE_ANODISATION.C_C16_IMPOSE, "C16 IMPOSE")
                            AffichageTexte MSHFGPrevisionnel, UnTexte
                        End If
                        
                    End If
                Next a

                '--- restitution des valeurs ligne, colonne ---
                .Redraw = True
                .Row = MemLigne
                .Col = MemColonne
                .FocusRect = flexFocusHeavy

            End With
        
        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Calcul du temps total de la gamme redresseur
' Entrées : CalculTempsTotalGammeRedresseur -> Le temps total de la gamme en secondes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CalculTempsTotalGammeRedresseur() As Long
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                            'pour les boucles FOR...NEXT

    '--- affectation ---
    CalculTempsTotalGammeRedresseur = 0
    
    '--- calcul du temps ---
    For a = MEBTempsPhases.LBound To MEBTempsPhases.UBound
        CalculTempsTotalGammeRedresseur = CalculTempsTotalGammeRedresseur + CTempsTexteEnSecondes(MEBTempsPhases(a).Text)
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion de la grille de recherche
' Entrées :  NumOnglet  -> Numéro de l'onglet fonction de l'énumération ONGLETS_CHARGEMENT_PREVISIONNEL
'                EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionGrilleRecherche(ByVal NumOnglet As ONGLETS_CHARGEMENT_PREVISIONNEL, _
                                                           ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    
    '--- affectation ---

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGGrilleRecherche(NumOnglet)
                
                .Visible = False                                                            'rendre la grille invisible
                '.ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'autorise le fractionnement de la grille (petite rectangle noir en bas à gauche)
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                
                If NumOnglet = O_CHARGEMENT Then
                    .HeadBackColor = COULEURS.ROUGE_5               'couleur de fond des entêtes
                Else
                    .HeadBackColor = COULEURS.BLEU_5                  'couleur de fond des entêtes
                End If
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.JAUNE_0              'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                
                .EvenRowStyle.BackColor = COULEURS.JAUNE_1      'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.CYAN_1         'couleur des lignes impaires
                
                .SelectedBackColor = COULEURS.ROUGE_3                'couleur de fond pour la sélection
                .SelectedForeColor = COULEURS.JAUNE_3                  'couleur de premier plan pour la sélection
                
                .HeadFont.Name = "Arial"
                With .Font
                    .Name = "MS Sans serif"
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.05
                
                .RecordSelectors = True                                                 'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = Me.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = Me.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = Me.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = False                                                  'interdire un nouvel enregistrement
                .AllowDelete = False                                                     'interdire la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire la sélection des colonnes
                .AllowColMove = False                                                 'interdire le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'autoriser la modification de l'épaisseur des lignes
                
                
                .DataView = dbgNormalView                                         'présentation normale de la grille
                
                .Visible = True                                                               'rendre la grille visible
            
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGGrilleRecherche(NumOnglet)
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_GAMME)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° de gamme"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_REF_GAMME)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Référence de la gamme"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NOM_GAMME)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom de la gamme"
                    .Width = EPAISSEUR_CARACTERE * 50
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

            End With

        Case Else

    End Select
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lance une recherche en fonction des critères
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LanceRechercheOuTri(ByVal NumOnglet As ONGLETS_CHARGEMENT_PREVISIONNEL)
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim IdxRecherchePar As Integer
    Dim RechercherPar As String, _
           CommencantPar As String, _
           Contenant  As String, _
           ChaineDeRecherche As String, _
           RequeteSQL As String, _
           Filtre1 As String, _
           Filtre2 As String
    
    '--- curseur de la souris ---
    SourisEnAttente True

    '--- affectation ---
    CommencantPar = TBCommencantPar(NumOnglet).Text
    Contenant = TBContenant(NumOnglet).Text
    IdxRecherchePar = Succ(CBRechercherPar(NumOnglet).ListIndex)
RechercherPar = Choose(IdxRecherchePar, _
                                              "NumGamme", _
                                              "RefGamme", _
                                              "NomGamme")
            
    '--- début de la requête ---
    RequeteSQL = "SELECT GammesAnodisation.* FROM GammesAnodisation "

    '--- modification pour le cas du numéro de la gamme d'anodisation ---
    Select Case IdxRecherchePar
        Case IDX_RECHERCHER_PAR.IDX_NUM_GAMME
            '--- cas du numéro de la gamme d'anodisation ---
            If CommencantPar <> "" Then
                CommencantPar = Right(FORMAT_NUM_GAMME_ANODISATION & CommencantPar, 6)
            End If
        Case Else
    End Select
    
    '--- filtres pour chaines de caractères ---
    Filtre1 = "(" & RechercherPar & " LIKE '" & CommencantPar & "%') "
    Filtre2 = "(" & RechercherPar & " LIKE '%" & Contenant & "%') "
    
    '--- construction du filtre ---
    If CommencantPar = "" And Contenant = "" Then
    ElseIf CommencantPar <> "" And Contenant = "" Then
        RequeteSQL = RequeteSQL & "WHERE " & Filtre1
    ElseIf CommencantPar = "" And Contenant <> "" Then
        RequeteSQL = RequeteSQL & "WHERE " & Filtre2
    Else
        RequeteSQL = RequeteSQL & "WHERE " & Filtre1 & " AND " & Filtre2
    End If
    
    '--- fin de la requête ---
    RequeteSQL = RequeteSQL & "ORDER BY " & RechercherPar
    Select Case IdxRecherchePar
        Case 1: RequeteSQL = RequeteSQL & ", DateCreationGamme DESC"                          'NumGamme
        Case 2: RequeteSQL = RequeteSQL & ", NumGamme"                                                  'RefGamme
        Case 3: RequeteSQL = RequeteSQL & ", NumGamme, DateCreationGamme DESC"    'NomGamme
        Case Else
    End Select

    'Debug.Print RequeteSQL
    With ADODCGammesAnodisation(NumOnglet)
        
        '--- application de la requête ---
        .Recordset.Cancel
        If .RecordSource <> RequeteSQL Then
            .RecordSource = RequeteSQL
            .Refresh
        Else
            .Recordset.Requery
        End If
        
        '--- message si fiche non trouvée ---
        With .Recordset
            If .EOF Or .BOF Then
                MessageErreur TITRE_MESSAGES, MESSAGE_121
            End If
        End With
    
    End With

    '--- curseur de la souris ---
    SourisEnAttente False
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- curseur de la souris ---
    SourisEnAttente False
    
End Sub



VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FGammesAnodisation 
   ClientHeight    =   13005
   ClientLeft      =   405
   ClientTop       =   3870
   ClientWidth     =   13395
   Icon            =   "FGammesAnodisation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   18981.91
   ScaleMode       =   0  'User
   ScaleWidth      =   1.67170e5
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
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
      Height          =   1215
      Left            =   240
      ScaleHeight     =   1185
      ScaleWidth      =   28185
      TabIndex        =   37
      Top             =   600
      Width           =   28215
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
         ItemData        =   "FGammesAnodisation.frx":014A
         Left            =   1680
         List            =   "FGammesAnodisation.frx":015A
         Style           =   2  'Dropdown List
         TabIndex        =   150
         Top             =   420
         Width           =   3495
      End
      Begin VB.OptionButton OBFormeGrille 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   " Change la pr�sentation de la grille "
         Top             =   420
         Width           =   375
      End
      Begin VB.OptionButton OBFormeGrille 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   60
         Value           =   -1  'True
         Width           =   375
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
         Height          =   735
         Left            =   5400
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":01A5
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   " Annule tris et recherches "
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CBRechercherSurGrille 
         BackColor       =   &H00E0E0E0&
         Caption         =   "GRILLE"
         DownPicture     =   "FGammesAnodisation.frx":0397
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
         Left            =   240
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":0A99
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   " Rechercher sur la grille "
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Height          =   735
         Left            =   11400
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":119B
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   " Lancer une recherche "
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1335
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
         Left            =   8460
         TabIndex        =   39
         Top             =   60
         Width           =   2655
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
         Left            =   8460
         TabIndex        =   38
         Top             =   480
         Width           =   2655
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGGrilleRecherche 
         Bindings        =   "FGammesAnodisation.frx":14DD
         Height          =   9915
         Left            =   240
         TabIndex        =   152
         Top             =   840
         Width           =   27675
         _ExtentX        =   48816
         _ExtentY        =   17489
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
         Columns(2).Caption=   "DateCreationGamme"
         Columns(2).DataField=   "DateCreationGamme"
         Columns(2).DataWidth=   19
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NomGamme"
         Columns(3).DataField=   "NomGamme"
         Columns(3).DataWidth=   50
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Matiere1"
         Columns(4).DataField=   "Matiere1"
         Columns(4).DataWidth=   30
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Matiere2"
         Columns(5).DataField=   "Matiere2"
         Columns(5).DataWidth=   30
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Matiere3"
         Columns(6).DataField=   "Matiere3"
         Columns(6).DataWidth=   30
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Matiere4"
         Columns(7).DataField=   "Matiere4"
         Columns(7).DataWidth=   30
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Matiere5"
         Columns(8).DataField=   "Matiere5"
         Columns(8).DataWidth=   30
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2566"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2434"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4154"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4022"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(2)._AlignLeft=0"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=4366"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=4233"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=4366"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=4233"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=4366"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=4233"
         Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(26)=   "Column(6).Width=4366"
         Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=4233"
         Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(30)=   "Column(7).Width=4366"
         Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=4233"
         Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(34)=   "Column(8).Width=4366"
         Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=4233"
         Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
      End
      Begin VB.Shape SFocusGrilleRecherche 
         BorderColor     =   &H000000FF&
         BorderWidth     =   4
         Height          =   330
         Left            =   120
         Top             =   1020
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
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
         Height          =   240
         Index           =   20
         Left            =   1740
         TabIndex        =   47
         Top             =   60
         Width           =   3360
      End
      Begin VB.Label LLibelles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commen�ant par"
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
         Height          =   240
         Index           =   13
         Left            =   6540
         TabIndex        =   46
         Top             =   120
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
         Height          =   240
         Index           =   12
         Left            =   6540
         TabIndex        =   45
         Top             =   540
         Width           =   1050
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FGammesAnodisation.frx":1502
      ScaleHeight     =   315
      ScaleWidth      =   13335
      TabIndex        =   8
      Top             =   0
      Width           =   13395
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "GAMME GEREE"
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
         Left            =   6780
         TabIndex        =   9
         Top             =   60
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
      ScaleWidth      =   13335
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   11910
      Width           =   13395
      Begin VB.Frame FNouveauNumGamme 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2640
         TabIndex        =   113
         Top             =   90
         Width           =   2415
         Begin VB.TextBox TBNouveauNumGamme 
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
            Left            =   630
            MaxLength       =   6
            TabIndex        =   114
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nouveau n� de gamme"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   49
            Left            =   0
            TabIndex        =   151
            Top             =   0
            Width           =   2415
         End
      End
      Begin VB.CommandButton CBCopieGammes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copie d'une gamme"
         DownPicture     =   "FGammesAnodisation.frx":25E44
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
         Left            =   120
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":26DE6
         Style           =   1  'Graphical
         TabIndex        =   112
         Top             =   105
         UseMaskColor    =   -1  'True
         Width           =   2355
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FGammesAnodisation.frx":27D88
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
         Left            =   27120
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":2848A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   " Quitter cette fen�tre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Timer TimerSimulationEntreeCharge 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   11400
         Top             =   180
      End
      Begin MSAdodcLib.Adodc ADODCGammesAnodisation 
         Height          =   435
         Left            =   21840
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualise&r"
         DownPicture     =   "FGammesAnodisation.frx":28B8C
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
         Left            =   18420
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":2928E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Actualiser les donn�es "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBNouveau 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Nouveau"
         DownPicture     =   "FGammesAnodisation.frx":29990
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
         Left            =   20160
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":2A092
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Nouvel enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FGammesAnodisation.frx":2A794
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
         Left            =   25380
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":2AE96
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FGammesAnodisation.frx":2B598
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
         Left            =   23640
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":2BC9A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Annuler les derni�res modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBSupprimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Supprimer"
         DownPicture     =   "FGammesAnodisation.frx":2C39C
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
         Left            =   14040
         MaskColor       =   &H00FF00FF&
         Picture         =   "FGammesAnodisation.frx":2CA9E
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   " Supprimer l'enregistrement en cours "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBVerifierCoherenceGamme 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&V�rifier la coh�rence"
         DownPicture     =   "FGammesAnodisation.frx":2D1A0
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
         Picture         =   "FGammesAnodisation.frx":2E142
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " V�rifie la coh�rence de la gamme "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2535
      End
      Begin MSComctlLib.ImageList ILOutilsGestionGrilles 
         Left            =   11940
         Top             =   180
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
               Picture         =   "FGammesAnodisation.frx":2F0E4
               Key             =   "supprimer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":304CE
               Key             =   "compacter"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":318B8
               Key             =   "inserer"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILGrillesDonnees 
         Left            =   12600
         Top             =   180
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
               Picture         =   "FGammesAnodisation.frx":32CA2
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":32EAE
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":330BA
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":332C6
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":334D2
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":336DE
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":338EA
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":33AF6
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":33D02
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":33F0E
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":3411A
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":34326
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":34532
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":3473E
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":3494A
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":34B56
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":34D62
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":34F66
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":3516A
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":3536E
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":35572
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":35776
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":3597A
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":35B7E
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":35D82
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FGammesAnodisation.frx":35F86
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   315
         Left            =   13200
         Top             =   240
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label LRenseignements 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   21840
         TabIndex        =   6
         Top             =   120
         Width           =   1455
      End
   End
   Begin C1SizerLibCtl.C1Tab CTOnglets 
      Height          =   10155
      Left            =   240
      TabIndex        =   11
      Top             =   2760
      Width           =   28215
      _cx             =   49768
      _cy             =   17912
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
      Caption         =   "Renseignements|D�tails de la gamme d'ANODISATION|Calculs par apprentissage"
      Align           =   0
      CurrTab         =   0
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
      Picture(0)      =   "FGammesAnodisation.frx":3618A
      Picture(1)      =   "FGammesAnodisation.frx":362E4
      Picture(2)      =   "FGammesAnodisation.frx":3643E
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   9
         Left            =   30060
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   21
         Top             =   495
         Width           =   28125
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   1
         Left            =   28860
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   20
         Top             =   495
         Width           =   28125
         Begin MSMask.MaskEdBox MEBEditionDetailsGammesAnodisation 
            Height          =   255
            Left            =   7920
            TabIndex        =   31
            Top             =   1140
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            ClipMode        =   1
            Appearance      =   0
            BackColor       =   12632319
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
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsGammesAnodisation 
            Height          =   6735
            Left            =   7680
            TabIndex        =   32
            Top             =   900
            Width           =   13455
            _ExtentX        =   23733
            _ExtentY        =   11880
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   12582912
            Rows            =   31
            Cols            =   6
            BackColorFixed  =   16576
            ForeColorFixed  =   16777215
            BackColorBkg    =   12648447
            GridColor       =   12632256
            GridColorUnpopulated=   -2147483644
            WordWrap        =   -1  'True
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
            ScrollBars      =   2
            Appearance      =   0
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
            Height          =   9135
            Left            =   21360
            TabIndex        =   115
            Top             =   180
            Visible         =   0   'False
            Width           =   6555
            Begin VB.PictureBox PBPhasesRedresseurs 
               BackColor       =   &H00C0E0FF&
               Height          =   3735
               Left            =   240
               ScaleHeight     =   3675
               ScaleWidth      =   6015
               TabIndex        =   116
               Top             =   2940
               Width           =   6075
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
                  TabIndex        =   118
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
                  Index           =   2
                  Left            =   2880
                  MaxLength       =   6
                  TabIndex        =   122
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
                  Index           =   3
                  Left            =   2880
                  MaxLength       =   6
                  TabIndex        =   128
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
                  Index           =   4
                  Left            =   2880
                  MaxLength       =   6
                  TabIndex        =   134
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
                  Index           =   1
                  Left            =   4440
                  MaxLength       =   6
                  TabIndex        =   119
                  Top             =   840
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
                  TabIndex        =   124
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
                  Index           =   3
                  Left            =   4440
                  MaxLength       =   6
                  TabIndex        =   130
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
                  Index           =   4
                  Left            =   4440
                  MaxLength       =   6
                  TabIndex        =   136
                  Top             =   2460
                  Width           =   855
               End
               Begin MSMask.MaskEdBox MEBTempsPhases 
                  Height          =   315
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   117
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
                  TabIndex        =   120
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
                  TabIndex        =   126
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
                  TabIndex        =   132
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
                  Index           =   46
                  Left            =   3840
                  TabIndex        =   145
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
                  Index           =   45
                  Left            =   3840
                  TabIndex        =   144
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
                  Index           =   44
                  Left            =   3840
                  TabIndex        =   143
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
                  Index           =   43
                  Left            =   3840
                  TabIndex        =   142
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
                  Index           =   42
                  Left            =   5400
                  TabIndex        =   141
                  Top             =   870
                  Width           =   195
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
                  Index           =   41
                  Left            =   2760
                  TabIndex        =   140
                  Top             =   360
                  Width           =   1335
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
                  Index           =   40
                  Left            =   4320
                  TabIndex        =   139
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
                  Index           =   39
                  Left            =   5400
                  TabIndex        =   138
                  Top             =   1410
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
                  Index           =   38
                  Left            =   5400
                  TabIndex        =   137
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
                  Index           =   37
                  Left            =   5400
                  TabIndex        =   135
                  Top             =   2490
                  Width           =   195
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
                  Index           =   36
                  Left            =   480
                  TabIndex        =   133
                  Top             =   840
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
                  Index           =   35
                  Left            =   480
                  TabIndex        =   131
                  Top             =   1380
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
                  Index           =   34
                  Left            =   480
                  TabIndex        =   129
                  Top             =   1920
                  Width           =   630
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
                  Index           =   33
                  Left            =   480
                  TabIndex        =   127
                  Top             =   2460
                  Width           =   630
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
                  Index           =   32
                  Left            =   1440
                  TabIndex        =   125
                  Top             =   360
                  Width           =   1095
               End
               Begin VB.Line LDecoration 
                  Index           =   12
                  X1              =   2640
                  X2              =   2640
                  Y1              =   720
                  Y2              =   2880
               End
               Begin VB.Line LDecoration 
                  Index           =   11
                  X1              =   1320
                  X2              =   1320
                  Y1              =   660
                  Y2              =   3420
               End
               Begin VB.Line LDecoration 
                  Index           =   10
                  X1              =   4200
                  X2              =   4200
                  Y1              =   720
                  Y2              =   2880
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
                  TabIndex        =   123
                  Top             =   3015
                  Width           =   855
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
                  Index           =   30
                  Left            =   480
                  TabIndex        =   121
                  Top             =   3000
                  Width           =   630
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
                  Index           =   1
                  Left            =   240
                  Shape           =   4  'Rounded Rectangle
                  Top             =   1260
                  Width           =   5535
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
                  Index           =   9
                  Left            =   1320
                  Shape           =   4  'Rounded Rectangle
                  Top             =   240
                  Width           =   1335
               End
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
               TabIndex        =   149
               Top             =   1260
               Width           =   915
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
               TabIndex        =   148
               Top             =   840
               Width           =   915
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
               Index           =   48
               Left            =   240
               TabIndex        =   147
               Top             =   420
               Width           =   2910
            End
            Begin VB.Image IPhasesAnodisation 
               Height          =   2010
               Left            =   240
               Picture         =   "FGammesAnodisation.frx":36598
               Top             =   720
               Width           =   2925
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
               Index           =   47
               Left            =   3150
               TabIndex        =   146
               Top             =   420
               Width           =   1170
            End
            Begin VB.Shape SDecoration 
               BorderWidth     =   2
               FillColor       =   &H00FFFFC0&
               FillStyle       =   0  'Solid
               Height          =   960
               Index           =   3
               Left            =   3150
               Top             =   735
               Width           =   1170
            End
         End
         Begin VB.Frame FTempsBainsEgouttages 
            Caption         =   " Le cumul des temps (POSTES + EGOUTTAGES) "
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
            Index           =   0
            Left            =   7680
            TabIndex        =   22
            Top             =   7800
            Width           =   13455
            Begin VB.Image Image1 
               Height          =   480
               Left            =   8940
               Picture         =   "FGammesAnodisation.frx":499A2
               Top             =   840
               Width           =   480
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "AVANT le POSTE PRINCIPAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   2
               Left            =   780
               TabIndex        =   30
               Top             =   360
               Width           =   2475
               WordWrap        =   -1  'True
            End
            Begin VB.Label LTempsAvantPostePrincipalSansPonts 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               Height          =   315
               Index           =   0
               Left            =   780
               TabIndex        =   29
               Top             =   960
               Width           =   2475
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "AU POSTE PRINCIPAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   1
               Left            =   3600
               TabIndex        =   28
               Top             =   360
               Width           =   2235
               WordWrap        =   -1  'True
            End
            Begin VB.Label LTempsPostePrincipalSansPonts 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
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
               Height          =   315
               Index           =   0
               Left            =   3480
               TabIndex        =   27
               Top             =   960
               Width           =   2475
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "APRES le POSTE PRINCIPAL"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   540
               Index           =   4
               Left            =   6120
               TabIndex        =   26
               Top             =   360
               Width           =   2445
               WordWrap        =   -1  'True
            End
            Begin VB.Label LTempsApresPostePrincipalSansPonts 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
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
               Height          =   315
               Index           =   0
               Left            =   6120
               TabIndex        =   25
               Top             =   960
               Width           =   2475
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "TOTAL de la GAMME"
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
               Index           =   5
               Left            =   10080
               TabIndex        =   24
               Top             =   480
               Width           =   2295
               WordWrap        =   -1  'True
            End
            Begin VB.Label LTempsTotalGammeSansPonts 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFC0&
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
               Height          =   315
               Index           =   0
               Left            =   9780
               TabIndex        =   23
               Top             =   960
               Width           =   2895
            End
         End
         Begin MSAdodcLib.Adodc ADODCZones 
            Height          =   375
            Left            =   300
            Top             =   8940
            Width           =   7095
            _ExtentX        =   12515
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
            Appearance      =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            Orientation     =   0
            Enabled         =   -1
            Connect         =   ""
            OLEDBString     =   ""
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   ""
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            _Version        =   393216
         End
         Begin MSDataGridLib.DataGrid DGZones 
            Bindings        =   "FGammesAnodisation.frx":49DE4
            Height          =   8355
            Left            =   300
            TabIndex        =   33
            Top             =   300
            Width           =   7095
            _ExtentX        =   12515
            _ExtentY        =   14737
            _Version        =   393216
            AllowUpdate     =   0   'False
            AllowArrows     =   -1  'True
            BackColor       =   16777215
            ForeColor       =   0
            HeadLines       =   2
            RowHeight       =   19
            TabAction       =   2
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   "CodeZone"
               Caption         =   "Code de la zone"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1036
                  SubFormatType   =   0
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   "LibelleZone"
               Caption         =   "Libell� de la zone"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1036
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               SizeMode        =   1
               ScrollGroup     =   0
               BeginProperty Column00 
                  Alignment       =   2
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   2025,071
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   3
                  Locked          =   -1  'True
                  ColumnWidth     =   4500,284
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.Toolbar TOBGestionGrilles 
            Height          =   405
            Left            =   7680
            TabIndex        =   75
            Top             =   300
            Width           =   13470
            _ExtentX        =   23760
            _ExtentY        =   714
            ButtonWidth     =   2514
            ButtonHeight    =   661
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Style           =   1
            ImageList       =   "ILOutilsGestionGrilles"
            HotImageList    =   "ILOutilsGestionGrilles"
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
                  Object.ToolTipText     =   " Ins�re une ligne dans une grille "
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Shape SFocusTableZones 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Height          =   8370
            Left            =   300
            Top             =   300
            Visible         =   0   'False
            Width           =   7110
         End
         Begin VB.Shape SFocusTableDetailsGammesAnodisation 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Height          =   6750
            Left            =   7680
            Top             =   900
            Visible         =   0   'False
            Width           =   13470
         End
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   0
         Left            =   45
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   19
         Top             =   495
         Width           =   28125
         Begin VB.Frame FCaracteristiques 
            Caption         =   " Mati�res concern�es "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7410
            Left            =   240
            TabIndex        =   36
            Top             =   1920
            Width           =   20535
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":49DFD
               Height          =   375
               Index           =   10
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4A4FF
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   5280
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":4AC01
               Height          =   375
               Index           =   9
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4B303
               Style           =   1  'Graphical
               TabIndex        =   84
               Top             =   4740
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":4BA05
               Height          =   375
               Index           =   8
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4C107
               Style           =   1  'Graphical
               TabIndex        =   83
               Top             =   4200
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":4C809
               Height          =   375
               Index           =   7
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4CF0B
               Style           =   1  'Graphical
               TabIndex        =   82
               Top             =   3660
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":4D60D
               Height          =   375
               Index           =   6
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4DD0F
               Style           =   1  'Graphical
               TabIndex        =   81
               Top             =   3120
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":4E411
               Height          =   375
               Index           =   5
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4EB13
               Style           =   1  'Graphical
               TabIndex        =   80
               Top             =   2580
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":4F215
               Height          =   375
               Index           =   4
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":4F917
               Style           =   1  'Graphical
               TabIndex        =   79
               Top             =   2040
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":50019
               Height          =   375
               Index           =   3
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":5071B
               Style           =   1  'Graphical
               TabIndex        =   78
               Top             =   1500
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":50E1D
               Height          =   375
               Index           =   2
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":5151F
               Style           =   1  'Graphical
               TabIndex        =   77
               Top             =   960
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.CommandButton CBSupprimerMatieres 
               DownPicture     =   "FGammesAnodisation.frx":51C21
               Height          =   375
               Index           =   1
               Left            =   5460
               MaskColor       =   &H00FF00FF&
               Picture         =   "FGammesAnodisation.frx":52323
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   420
               UseMaskColor    =   -1  'True
               Width           =   435
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere10"
               DataSource      =   "ADODCGammesAnodisation"
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
               Index           =   10
               Left            =   780
               TabIndex        =   74
               Top             =   5280
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere9"
               DataSource      =   "ADODCGammesAnodisation"
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
               Index           =   9
               Left            =   780
               TabIndex        =   63
               Top             =   4740
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere8"
               DataSource      =   "ADODCGammesAnodisation"
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
               Index           =   8
               Left            =   780
               TabIndex        =   62
               Top             =   4200
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere7"
               DataSource      =   "ADODCGammesAnodisation"
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
               Index           =   7
               Left            =   780
               TabIndex        =   61
               Top             =   3660
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere6"
               DataSource      =   "ADODCGammesAnodisation"
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
               Index           =   6
               Left            =   780
               TabIndex        =   60
               Top             =   3120
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere5"
               DataSource      =   "ADODCGammesAnodisation"
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
               Index           =   5
               Left            =   780
               TabIndex        =   59
               Top             =   2580
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere4"
               DataSource      =   "ADODCGammesAnodisation"
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
               Left            =   780
               TabIndex        =   58
               Top             =   2040
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere3"
               DataSource      =   "ADODCGammesAnodisation"
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
               Left            =   780
               TabIndex        =   57
               Top             =   1500
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere2"
               DataSource      =   "ADODCGammesAnodisation"
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
               Left            =   780
               TabIndex        =   56
               Top             =   960
               Width           =   4575
            End
            Begin VB.TextBox TBMatieres 
               DataField       =   "Matiere1"
               DataSource      =   "ADODCGammesAnodisation"
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
               Left            =   780
               TabIndex        =   55
               Top             =   420
               Width           =   4575
            End
            Begin MSAdodcLib.Adodc ADODCMatieres 
               Height          =   375
               Left            =   6120
               Top             =   6780
               Width           =   14175
               _ExtentX        =   25003
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
               Appearance      =   0
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               Orientation     =   0
               Enabled         =   -1
               Connect         =   ""
               OLEDBString     =   ""
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   ""
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               _Version        =   393216
            End
            Begin MSDataGridLib.DataGrid DGMatieres 
               Bindings        =   "FGammesAnodisation.frx":52A25
               Height          =   6135
               Left            =   6120
               TabIndex        =   86
               Top             =   420
               Width           =   14175
               _ExtentX        =   25003
               _ExtentY        =   10821
               _Version        =   393216
               AllowUpdate     =   0   'False
               AllowArrows     =   -1  'True
               BackColor       =   16777152
               ForeColor       =   0
               HeadLines       =   2
               RowHeight       =   19
               TabAction       =   2
               FormatLocked    =   -1  'True
               BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnCount     =   4
               BeginProperty Column00 
                  DataField       =   "OrdrePourAffichage"
                  Caption         =   "Ordre"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1036
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column01 
                  DataField       =   "Matiere"
                  Caption         =   "Matiere"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1036
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column02 
                  DataField       =   "TypeMatiere"
                  Caption         =   "Type"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1036
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               BeginProperty Column03 
                  DataField       =   "CompositionMatiere"
                  Caption         =   "Composition"
                  BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                     Type            =   0
                     Format          =   ""
                     HaveTrueFalseNull=   0
                     FirstDayOfWeek  =   0
                     FirstWeekOfYear =   0
                     LCID            =   1036
                     SubFormatType   =   0
                  EndProperty
               EndProperty
               SplitCount      =   1
               BeginProperty Split0 
                  MarqueeStyle    =   3
                  SizeMode        =   1
                  ScrollGroup     =   0
                  Size            =   0
                  BeginProperty Column00 
                     Locked          =   -1  'True
                     ColumnWidth     =   780,095
                  EndProperty
                  BeginProperty Column01 
                     Locked          =   -1  'True
                     ColumnWidth     =   2775,118
                  EndProperty
                  BeginProperty Column02 
                     Locked          =   -1  'True
                     ColumnWidth     =   2775,118
                  EndProperty
                  BeginProperty Column03 
                     Locked          =   -1  'True
                     ColumnWidth     =   7274,835
                  EndProperty
               EndProperty
            End
            Begin VB.Shape SFocusTableMatieres 
               BorderColor     =   &H000000FF&
               BorderWidth     =   4
               Height          =   6150
               Left            =   6120
               Top             =   420
               Visible         =   0   'False
               Width           =   14190
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "10"
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
               Left            =   240
               TabIndex        =   73
               Top             =   5280
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "9"
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
               Left            =   240
               TabIndex        =   72
               Top             =   4740
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "8"
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
               Left            =   240
               TabIndex        =   71
               Top             =   4200
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "7"
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
               Left            =   240
               TabIndex        =   70
               Top             =   3660
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "6"
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
               Left            =   240
               TabIndex        =   69
               Top             =   3120
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5"
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
               Left            =   240
               TabIndex        =   68
               Top             =   2580
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4"
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
               Left            =   240
               TabIndex        =   67
               Top             =   2040
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3"
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
               Left            =   240
               TabIndex        =   66
               Top             =   1500
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2"
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
               Left            =   240
               TabIndex        =   65
               Top             =   960
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1"
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
               Left            =   240
               TabIndex        =   64
               Top             =   420
               Width           =   435
            End
         End
         Begin VB.Frame FDesignation 
            Caption         =   " D�signation "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   240
            TabIndex        =   34
            Top             =   180
            Width           =   20535
            Begin VB.TextBox TBDesignation 
               BackColor       =   &H00FFFFFF&
               DataField       =   "Designation"
               DataSource      =   "ADODCGammesAnodisation"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   180
               MaxLength       =   255
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   35
               Top             =   300
               Width           =   20175
            End
         End
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   4
         Left            =   29760
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   18
         Top             =   495
         Width           =   28125
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   5
         Left            =   30360
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   17
         Top             =   495
         Width           =   28125
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   6
         Left            =   30660
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   16
         Top             =   495
         Width           =   28125
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   7
         Left            =   30960
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   15
         Top             =   495
         Width           =   28125
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   8
         Left            =   31260
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   14
         Top             =   495
         Width           =   28125
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   2
         Left            =   29160
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   13
         Top             =   495
         Width           =   28125
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "SIMULATION de l'ENTREE d'une CHARGE AVEC CETTE GAMME"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   31
            Left            =   10440
            TabIndex        =   111
            Top             =   720
            Width           =   9975
         End
         Begin VB.Label LSimulationEntreeCharge 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   0
            Left            =   10500
            TabIndex        =   110
            Top             =   1260
            Width           =   9855
         End
         Begin VB.Label LSimulationEntreeCharge 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   10500
            TabIndex        =   109
            Top             =   1620
            Width           =   9855
         End
         Begin VB.Label LTempsTotalGammeAvecPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Height          =   315
            Left            =   7980
            TabIndex        =   100
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label LTempsApresPostePrincipalAvecPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Height          =   315
            Left            =   7980
            TabIndex        =   101
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label LTempsTotalGammeSansPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
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
            Height          =   315
            Index           =   1
            Left            =   3780
            TabIndex        =   93
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label LTempsApresPostePrincipalSansPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Height          =   315
            Index           =   1
            Left            =   3780
            TabIndex        =   94
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label LTempsPostePrincipalSansPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Height          =   315
            Index           =   1
            Left            =   3780
            TabIndex        =   95
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label LTempsPostePrincipalAvecPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Height          =   315
            Left            =   7980
            TabIndex        =   108
            Top             =   3120
            Width           =   1815
         End
         Begin VB.Label LTempsAvantPostePrincipalAvecPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Height          =   315
            Left            =   7980
            TabIndex        =   102
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label LTempsMouvementsAvantPostePrincipal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Height          =   315
            Left            =   5880
            TabIndex        =   104
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label LTempsAvantPostePrincipalSansPonts 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Height          =   315
            Index           =   1
            Left            =   3780
            TabIndex        =   96
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ATTENTION - Ces r�sultats th�oriques peuvent varier de quelques dizaines de secondes avec la r�alit�"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   345
            Index           =   7
            Left            =   180
            TabIndex        =   89
            Top             =   180
            Width           =   27735
            WordWrap        =   -1  'True
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Dot
            Index           =   9
            X1              =   5340
            X2              =   8520
            Y1              =   3270
            Y2              =   3270
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Dot
            Index           =   8
            X1              =   5250
            X2              =   8430
            Y1              =   3255
            Y2              =   3255
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   4
            X1              =   7680
            X2              =   7980
            Y1              =   3780
            Y2              =   3780
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   5
            X1              =   7680
            X2              =   7980
            Y1              =   4260
            Y2              =   4260
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   3
            X1              =   7680
            X2              =   7980
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   2
            X1              =   5580
            X2              =   5880
            Y1              =   4260
            Y2              =   4260
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   1
            X1              =   5580
            X2              =   5880
            Y1              =   3780
            Y2              =   3780
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Index           =   0
            X1              =   5580
            X2              =   6000
            Y1              =   2820
            Y2              =   2820
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Dot
            Index           =   6
            X1              =   5520
            X2              =   8040
            Y1              =   3285
            Y2              =   3285
         End
         Begin VB.Line LDecoration 
            BorderColor     =   &H000000FF&
            BorderStyle     =   3  'Dot
            Index           =   7
            X1              =   5340
            X2              =   8520
            Y1              =   3240
            Y2              =   3240
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " MOUVEMENTS des PONTS"
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
            Height          =   555
            Index           =   19
            Left            =   5880
            TabIndex        =   107
            Top             =   1860
            Width           =   1815
         End
         Begin VB.Label LTempsMouvementsApresPostePrincipal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Height          =   315
            Left            =   5880
            TabIndex        =   106
            Top             =   3600
            Width           =   1815
         End
         Begin VB.Label LTempsTotalMouvements 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0FF&
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
            Height          =   315
            Left            =   5880
            TabIndex        =   105
            Top             =   4080
            Width           =   1815
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AVANT LE POSTE PRINCIPAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   300
            TabIndex        =   103
            Top             =   2700
            Width           =   3315
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "AU POSTE PRINCIPAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   15
            Left            =   1080
            TabIndex        =   99
            Top             =   3180
            Width           =   2535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "APRES LE POSTE PRINCIPAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   17
            Left            =   420
            TabIndex        =   98
            Top             =   3660
            Width           =   3195
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   18
            Left            =   2580
            TabIndex        =   97
            Top             =   4140
            Width           =   1035
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "POSTES + EGOUTTAGES"
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
            Height          =   555
            Index           =   10
            Left            =   3780
            TabIndex        =   92
            Top             =   1860
            Width           =   1815
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "POSTES + EGOUTTAGES + MOUVEMENTS des PONTS"
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
            Height          =   1035
            Index           =   8
            Left            =   7980
            TabIndex        =   91
            Top             =   1380
            Width           =   1815
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "CALCULS PAR APPRENTISSAGE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   22
            Left            =   600
            TabIndex        =   90
            Top             =   720
            Width           =   9135
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   3615
            Index           =   4
            Left            =   180
            Shape           =   4  'Rounded Rectangle
            Top             =   1080
            Width           =   9975
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   1035
            Index           =   0
            Left            =   10320
            Shape           =   4  'Rounded Rectangle
            Top             =   1080
            Width           =   10215
         End
      End
      Begin VB.PictureBox PBOnglets 
         Height          =   9615
         Index           =   3
         Left            =   29460
         ScaleHeight     =   9555
         ScaleWidth      =   28065
         TabIndex        =   12
         Top             =   495
         Width           =   28125
      End
   End
   Begin VB.PictureBox PBCommuns 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      ScaleHeight     =   585
      ScaleWidth      =   28185
      TabIndex        =   48
      Top             =   1920
      Width           =   28215
      Begin VB.TextBox TBRefGamme 
         BackColor       =   &H00FFFFFF&
         DataField       =   "RefGamme"
         DataSource      =   "ADODCGammesAnodisation"
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
         Left            =   6480
         TabIndex        =   88
         Top             =   120
         Width           =   6975
      End
      Begin VB.TextBox TBNomGamme 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NomGamme"
         DataSource      =   "ADODCGammesAnodisation"
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
         Left            =   15420
         TabIndex        =   50
         Top             =   120
         Width           =   12615
      End
      Begin VB.TextBox TBNumGamme 
         BackColor       =   &H00C0FFFF&
         DataField       =   "NumGamme"
         DataSource      =   "ADODCGammesAnodisation"
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
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   49
         Top             =   120
         Width           =   1155
      End
      Begin MSMask.MaskEdBox MBDateGamme 
         DataField       =   "DateCreationGamme"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   3
         EndProperty
         DataSource      =   "ADODCGammesAnodisation"
         Height          =   315
         Left            =   2700
         TabIndex        =   51
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LLibelles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R�f�rence de la gamme"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   29
         Left            =   4140
         TabIndex        =   87
         Top             =   180
         Width           =   2235
         WordWrap        =   -1  'True
      End
      Begin VB.Label LLibelles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "du"
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
         Index           =   3
         Left            =   2340
         TabIndex        =   54
         Top             =   180
         Width           =   225
      End
      Begin VB.Label LLibelles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gamme"
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
         Index           =   0
         Left            =   180
         TabIndex        =   53
         Top             =   180
         Width           =   720
         WordWrap        =   -1  'True
      End
      Begin VB.Label LLibelles 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nom de la gamme"
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
         Index           =   9
         Left            =   13560
         TabIndex        =   52
         Top             =   180
         Width           =   1755
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "FGammesAnodisation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : Fen�tre g�rant les gammes d'anodisation
' Nom                    : FGammesAnodisation.frm
' Date de cr�ation : 06/10/2010
' D�tails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z
    
'--- constantes priv�es ---
Private Const NBR_COLONNES_DETAILS_GAMMES_PRODUCTION  As Integer = 6
Private Const TITRE_FENETRE As String = "GAMMES D'ANODISATION"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- �num�rations priv�es ---
Private Enum ONGLETS
    O_RENSEIGNEMENTS = 0
    O_GAMME_ANODISATION = 1
    O_CALCULS_PAR_APPRENTISSAGE = 2
End Enum

Private Enum IDX_RECHERCHER_PAR
    IDX_NUM_GAMME = 1
    IDX_REF_GAMME = 2
    IDX_DATE_CREATION_GAMME = 3
    IDX_NOM_GAMME = 4
End Enum

Private Enum COLONNES_GRILLE_RECHERCHE
    C_NUM_GAMME = 0
    C_REF_GAMME = 1
    C_DATE_CREATION_GAMME = 2
    C_NOM_GAMME = 3
    C_MATIERE_1 = 4
    C_MATIERE_2 = 5
    C_MATIERE_3 = 6
    C_MATIERE_4 = 7
    C_MATIERE_5 = 8
End Enum

Private Enum COLONNES_DETAILS_GAMMES_PRODUCTION
    C_NUM_LIGNES = 0
    C_CODE_ZONE = 1
    C_LIBELLE_ZONE = 2
    C_TEMPS_AU_POSTE_TEXTE = 3
    C_TEMPS_ALERTE_TEXTE = 4
    C_TEMPS_EGOUTTAGE_TEXTE = 5
    C_PONT = 6
End Enum

'--- types priv�s ---
Private Type ImgDetailsGammesProduction
    
    NumZone As Integer                              'N� de la zone
    Codezone As String                               'Code de la zone
    LibelleZone As String                            'Libell� de la zone
    
    TempsAuPosteTexte As String              'Temps au poste en texte au format HH:MM:SS
    TempsAlerteTexte As String                  'Temps d'alerte pour les colorants en texte au format HH:MM:SS
    TempsEgouttageTexte As String           'Temps d'�gouttage en texte au format MM:SS
    
    TempsAuPosteSecondes As Long         'Temps au poste en secondes
    TempsAlerteSecondes As Long             'Temps d'alerte en secondes
    TempsEgouttageSecondes As Integer   'Temps d'�gouttage en secondes

End Type

'--- variables priv�es ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                                            'pour interdire certains �v�nements
Private LigneDepartDeplacement As Integer                                       'ligne de d�part en cas de d�placement d'un d�tail
Private LigneArriveeDeplacement As Integer                                       'ligne de d'arriv�e en cas de d�placement d'un d�tail
Private MemDernierBouton As Long                                                    'm�moire du dernier bouton

Private TempsAvantPostePrincipalSansPontsSecondes As Long       'temps avant le poste principal sans les ponts en secondes
Private TempsPostePrincipalSansPontsSecondes As Long                'temps au poste principal sans les ponts en secondes
Private TempsApresPostePrincipalSansPontsSecondes As Long       'temps apr�s le poste principal sans les ponts en secondes
Private TempsTotalPostesSansPontsSecondes As Long                     'temps total des postes sans les ponts en secondes
Private TempsTotalEgouttagesSansPontsSecondes As Long              'temps total des �gouttages sans les ponts en secondes
Private TempsTotalGammeSansPontsSecondes As Long                    'temps total de la gamme sans les ponts en secondes

Private TempsMouvementsAvantPostePrincipalSecondes As Long     'temps des mouvements avant le poste principal en secondes
Private TempsAvantPostePrincipalAvecPontsSecondes As Long         'temps avant le poste principal avec les ponts en secondes
Private TempsPostePrincipalAvecPontsSecondes As Long                  'temps au poste d'anodisation avec les ponts en secondes
Private TempsMouvementsApresPostePrincipalSecondes As Long     'temps des mouvements apr�s le poste principal en secondes
Private TempsApresPostePrincipalAvecPontsSecondes As Long         'temps apr�s le poste principal avec les ponts en secondes
Private TempsTotalPostesAvecPontsSecondes As Long                      'temps total des postes avec les ponts en secondes
Private TempsTotalEgouttagesAvecPontsSecondes As Long               'temps total des �gouttages avec les ponts en secondes
Private TempsTotalMouvementsSecondes As Long                             'temps total des mouvements en secondes
Private TempsTotalGammeAvecPontsSecondes As Long                     'temps total de la gamme avec les ponts en secondes

Private TempsAvantPostePrincipalSansPontsTexte As String              'temps avant le poste principal sans les ponts en texte au format HH:MM:SS
Private TempsPostePrincipalSansPontsTexte As String                       'temps au poste principal sans les ponts en texte au format HH:MM:SS
Private TempsApresPostePrincipalSansPontsTexte As String              'temps apr�s poste principal sans les ponts en texte au format HH:MM:SS
Private TempsTotalPostesSansPontsTexte As String                           'temps total des postes sans les ponts en texte au format HH:MM:SS
Private TempsTotalEgouttagesSansPontsTexte As String                    'temps total des �gouttages sans les ponts en texte au format HH:MM:SS
Private TempsTotalGammeSansPontsTexte As String                          'temps total de la gamme sans les ponts en texte au format HH:MM:SS
    
Private TempsMouvementsAvantPostePrincipalTexte As String          'temps des mouvements avant le poste principal au format HH:MM:SS
Private TempsAvantPostePrincipalAvecPontsTexte As String              'temps avant le poste principal avec les ponts au format HH:MM:SS
Private TempsAnodisationAvecPontsTexte As String                            'temps au poste d'anodisation avec les ponts au format HH:MM:SS
Private TempsMouvementsApresPostePrincipalTexte As String          'temps des mouvements apr�s le poste principal au format HH:MM:SS
Private TempsApresPostePrincipalAvecPontsTexte As String              'temps apr�s le poste principal avec les ponts au format HH:MM:SS
Private TempsTotalPostesAvecPontsTexte As String                            'temps total des postes avec les ponts au format HH:MM:SS
Private TempsTotalEgouttagesAvecPontsTexte As String                     'temps total des �gouttages avec les ponts au format HH:MM:SS
Private TempsTotalMouvementsTexte As String                                   'temps total des mouvements au format HH:MM:SS
Private TempsTotalGammeAvecPontsTexte As String                          'temps total de la gamme avec les ponts au format HH:MM:SS

Private ModeUouIEnCours As MODES_U_OU_I                                    'mode U ou I en cours

'--- remarque ---
' par d�finition le temps d'anodisation sans les mouvements est identique � celui avec les mouvements

'--- tableaux priv�s ---
Private TDetailsGammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesProduction

'--- variables publiques ---
Public RechercherSurGrille As Boolean          'publique pour le copier / coller
Public NumFenetre As Long                             'num�ro de la fen�tre lorsqu'elle devient active

Sub Form_Initialize()

   
    

End Sub

Private Sub Form_Load()
    
            
     With ADODCGammesAnodisation
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .RecordSource = "SELECT GammesAnodisation.* From GammesAnodisation ORDER BY NumGamme"
        .Refresh
        
       
    End With
    With ADODCMatieres
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .RecordSource = "SELECT OrdrePourAffichage, Matiere, TypeMatiere,   CompositionMatiere From Matieres ORDER BY OrdrePourAffichage"
        .Refresh
        
    End With
    
    With ADODCZones
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .RecordSource = " SELECT Zones.* From Zones ORDER BY NumZone"

        .Refresh
        
    End With
  
    
End Sub


Private Sub ADODCMatieres_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
    With pRecordset
        If .BOF = False And .EOF = False Then
            ADODCMatieres.Caption = .Fields("Matiere") & Space(10) & .Fields("TypeMatiere") & Space(10) & .Fields("compositionMatiere")
        End If
    End With

End Sub

Private Sub ADODCZones_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
    With pRecordset
        If .BOF = False And .EOF = False Then
            ADODCZones.Caption = .Fields("LibelleZone")
        End If
    End With

End Sub

Private Sub ADODCGammesAnodisation_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer                    'pour les boucles FOR...NEXT
    
    With pRecordset
        
        If .BOF = False And .EOF = False Then
        
            '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
            Select Case MemDernierBouton
                
                Case ETATS_BOUTONS.E_AVANT_NOUVEAU, ETATS_BOUTONS.E_APRES_NOUVEAU
                    Me.Caption = TITRE_FENETRE & " - "
                    LRenseignements.Caption = "-"
                
                Case Else
                    '--- interdire les �v�nements ---
                    InterdireEvenements = True
                    
                    '--- affichage de la partie redresseur ---
                    Call LModeUouI_Click(pRecordset.Fields("ModeUouI").value)
                    
                    '--- phases ---
                    For a = PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4
                        MEBTempsPhases(a).Text = Right(CTemps2(pRecordset.Fields("TempsPhase" & a).value), 7)
                        TBTensionsPhases(a).Text = Format(pRecordset.Fields("UPhase" & a).value, FORMAT_TENSION_1_DECIMALE)
                        TBIntensitesPhases(a).Text = Format(pRecordset.Fields("IPhase" & a).value, FORMAT_INTENSITE_ENTIER)
                    Next a
                    
                    '--- calcul du temps total du cycle du redresseur ---
                    LTempsTotalGammeRedresseur.Caption = Right(CTemps2(CalculTempsTotalGammeRedresseur()), 7)
                    
                    '--- autoriser les �v�nements ---
                    InterdireEvenements = False
                    
                    '--- affichage de la position ---
                    If IsError(pRecordset("NumGamme")) = False Then
                        Me.Caption = TITRE_FENETRE & " - " & pRecordset("NumGamme")
                        LRenseignements.Caption = .AbsolutePosition & "/" & .RecordCount
                    End If
            
            End Select
            
        Else
            
            '--- si fiche inexistante affichage d'un tiret ---
            Me.Caption = TITRE_FENETRE
            LRenseignements.Caption = "-"
        
        End If
    
        '--- affichage des renseignements de la fenetre ---
        LRenseignementsFenetre.Caption = Me.Caption
    
    End With
    
    '--- chargement des d�tails ---
    If PremiereActivation = True And RechercherSurGrille = False Then
        LectureDetailsGammesAnodisation
    End If

End Sub

Private Sub CBActualiser_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_ACTUALISER
    
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- marquage ---
    MarqueEnregistrement True
    
    '--- actualisation ---
    ADODCGammesAnodisation.Refresh
    TDBGGrilleRecherche.Refresh
    ADODCMatieres.Refresh
    ADODCZones.Refresh
    
    '--- restitution ---
    MarqueEnregistrement False
    
    '--- curseur de la souris ---
    SourisEnAttente False

    '--- gestion des boutons ---
    GestionBoutons E_APRES_ACTUALISER
    
End Sub

Private Sub CBActualiser_GotFocus()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
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

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_ANNULER
        
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- annuler ---
    ADODCGammesAnodisation.Recordset.CancelUpdate
    
    '--- restitution ---
    MarqueEnregistrement False
    
    '--- curseur de la souris ---
    SourisEnAttente False
    
    '--- gestion des boutons ---
    GestionBoutons E_APRES_ANNULER
    
End Sub

Private Sub CBAnnuler_GotFocus()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
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

Private Sub CBCopieGammes_Click()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim NumGammeACopier As String                   'num�ro de la gamme � copier
    Dim NouveauNumGamme As String                 'nouveau num�ro de gamme

    '--- affectation du num�ro de gamme � copier et du nouveau num�ro de gamme ---
    NumGammeACopier = TBNumGamme.Text
    NouveauNumGamme = TBNouveauNumGamme.Text
    
    If NumGammeACopier <> "" And NouveauNumGamme <> "" Then

        If ExistenceGammesAnodisation(NouveauNumGamme) = TROUVE Then
            
            '--- messages relatifs aux gammes d'anodisation ---
            If AppelFenetre(F_MESSAGE, _
                                     TITRE_MESSAGES, _
                                     vbCrLf & vbCrLf & "c|Cette gamme existe d�j�" & vbCrLf & vbCrLf & "c|Vouez-vous la remplacer ?", _
                                     TYPES_MESSAGES.T_AVERTISSEMENT, _
                                     TYPES_BOUTONS.T_OUI_NON, _
                                     EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then

            
                '--- suppression des d�tails et de la gamme d'anodisation ---
                SuppressionGammesAnodisation NouveauNumGamme
                SuppressionDetailsGammesAnodisation NouveauNumGamme
            
                '--- copie de la gamme d'anodisation ---
                CopieGammeAnodisation NumGammeACopier:=NumGammeACopier, _
                                                          NouveauNumGamme:=NouveauNumGamme
            
                '--- actualisation ---
                Call CBActualiser_Click
            
                '--- effacement du champ du num�ro de gamme ---
                TBNouveauNumGamme.Text = ""
            
            End If
            
        Else
            
            '--- copie de la gamme d'anodisation ---
            CopieGammeAnodisation NumGammeACopier:=NumGammeACopier, _
                                                      NouveauNumGamme:=NouveauNumGamme
            
            '--- actualisation ---
            Call CBActualiser_Click
            
            '--- effacement du champ du num�ro de gamme ---
            TBNouveauNumGamme.Text = ""

        End If

    End If

End Sub

Private Sub CBLancerRecherche_Click()
    On Error Resume Next
    LanceRechercheOuTri
End Sub

Private Sub CBNouveau_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_NOUVEAU
   
    '--- marquage ---
    MarqueEnregistrement True
    
    '--- nouvel enregistrement ---
    ADODCGammesAnodisation.Recordset.AddNew

    '--- vidage des grilles ---
    GestionDetailsGammesAnodisation GG_VIDAGE
    GestionDetailsGammesAnodisation GG_AFFICHAGE
    
    '--- initialisation et affichage des temps de gamme ---
    AffichageTempsGamme
    AffichageCalculsParApprentissage
    
    '--- initialise les champs de la partie redresseur ---
    InitialisationChampsRedresseur
    
    '--- valeurs de champs ---
    With ADODCGammesAnodisation
        .Recordset(MBDateGamme.DataField) = Format(Now, "dd/mm/yyyy")
    End With
    CTOnglets.CurrTab = ONGLETS.O_RENSEIGNEMENTS
    
    '--- gestion des boutons ---
    GestionBoutons E_APRES_NOUVEAU
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
  
End Sub

Private Sub CBNouveau_GotFocus()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBNouveau_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
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
                CBValider_Click
                DechargeFenetre
            Case vbNo
                CBAnnuler_Click
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

    '--- d�placement du focus sur le bouton ---
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

Private Sub CBRaz_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- vidage des champs / lancement de la requ�te ---
    With TBCommencantPar
        .Text = ""
        .Refresh
        .SetFocus
    End With
    With TBContenant
        .Text = ""
        .Refresh
    End With
    LanceRechercheOuTri

End Sub

Private Sub CBRechercherPar_Click()
    On Error Resume Next
    If PremiereActivation = True Then
        DoEvents
        CBRaz_Click
    End If
End Sub

Private Sub CBRechercherSurGrille_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    If CBRechercherSurGrille.Enabled = True Then

        '--- affectation ---
        RechercherSurGrille = Not (RechercherSurGrille)
                
        '--- affichage ---
        AfficheGrilleRecherche
        
        '--- lancer la lecture des d�tails ---
        If PremiereActivation = True And RechercherSurGrille = False Then
            LectureDetailsGammesAnodisation
        End If

    End If

End Sub

Private Sub CBSupprimer_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_SUPPRIMER
    
    '--- demande de confirmation ---
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                            MESSAGE_2, _
                            TYPES_MESSAGES.T_AVERTISSEMENT, _
                            TYPES_BOUTONS.T_OUI_NON, _
                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
    
        '--- curseur de la souris ---
        SourisEnAttente True
        
        '--- suppression des d�tails ---
        SuppressionDetailsGammesAnodisation TBNumGamme.Text
        
        '--- effacement de l'enregistrement ---
        With ADODCGammesAnodisation.Recordset
            .Delete adAffectCurrent
            .UpdateBatch adAffectAllChapters
        End With
        
        '--- rafraichissement de la grille ---
        TDBGGrilleRecherche.Refresh
        
        '--- curseur de la souris ---
        SourisEnAttente False
    
    End If
    
    '--- gestion des boutons ---
    GestionBoutons E_APRES_SUPPRIMER

    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
    
End Sub

Private Sub CBSupprimer_GotFocus()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBSupprimer_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBSupprimerMatieres_Click(Index As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer                                                                                      'pour les boucles FOR...NEXT
    Dim Cpt As Integer                                                                                   'repr�sente un compteur
    Dim TMatieres(1 To NBR_MATIERES_MAXI_PAR_GAMME) As String   'tableau contenant les mati�res

    '--- effacement du champ ---
    With TBMatieres(Index)
        If .Text <> "" Then
            ADODCGammesAnodisation.Recordset("Matiere" & Index).value = ""
            GestionBoutons E_MODIFICATION_EN_COURS
        End If
    End With

    '--- m�morisation de champs ---
    For a = LBound(TMatieres()) To UBound(TMatieres())
        TMatieres(a) = TBMatieres(a).Text
    Next a

    '--- d�calage des champs ---
    Cpt = 1
    For a = LBound(TMatieres()) To UBound(TMatieres())
        If TMatieres(a) <> "" Then
            ADODCGammesAnodisation.Recordset("Matiere" & Cpt).value = TMatieres(a)
            Inc Cpt
        End If
    Next a

    '--- vidage du reste des champs ---
    For a = Cpt To NBR_MATIERES_MAXI_PAR_GAMME
        ADODCGammesAnodisation.Recordset("Matiere" & a).value = ""
    Next a

End Sub

Private Sub CBValider_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim PassageAnodisation As Boolean                           'indique un passage dans un des bains d'anodisation
    Dim PassageSpectro As Boolean                                  'indique un passage dans le bain de spectrocoloration
    Dim PassageOr As Boolean                                           'indique un passage dans le bain d'or
    Dim PassageNoir As Boolean                                        'indique un passage dans le bain de noir
            
    Dim a As Integer                    'pour les boucles FOR...NEXT
    Dim DernierEtat As Integer
    
    Dim NumGamme As String
    
    '--- gestion des boutons ---
    DernierEtat = MemDernierBouton
    GestionBoutons E_AVANT_VALIDER
    
    '--- curseur de la souris ---
    SourisEnAttente True
    
    
    'Call Log("CBValider_Click", "CBValider_Click")
    
    
    '--- contr�le sur la cl� primaire ---
    Select Case DernierEtat
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            NumGamme = ProchainNumGamme()
            ADODCGammesAnodisation.Recordset(TBNumGamme.DataField) = NumGamme
        Case Else
            NumGamme = Trim(TBNumGamme.Text)
    End Select
    
    '--- suppression et r�enregistrement des d�tails ---
    SuppressionDetailsGammesAnodisation NumGamme
    EnregistrementDetailsGammesAnodisation NumGamme
    
    '--- enregistrement de tous les temps ---
    With ADODCGammesAnodisation.Recordset
        
        .Fields("TempsAvantPostePrincipalSecondes").value = TempsAvantPostePrincipalSansPontsSecondes
        .Fields("TempsPostePrincipalSecondes").value = TempsPostePrincipalSansPontsSecondes
        .Fields("TempsApresPostePrincipalSecondes").value = TempsApresPostePrincipalSansPontsSecondes
        .Fields("TempsTotalPostesSecondes").value = TempsTotalPostesSansPontsSecondes
        .Fields("TempsTotalEgouttagesSecondes").value = TempsTotalEgouttagesSansPontsSecondes
        .Fields("TempsTotalGammeSecondes").value = TempsTotalGammeSansPontsSecondes
        .Fields("TempsAvantPostePrincipalTexte").value = TempsAvantPostePrincipalSansPontsTexte
        .Fields("TempsPostePrincipalTexte").value = TempsPostePrincipalSansPontsTexte
        .Fields("TempsApresPostePrincipalTexte").value = TempsApresPostePrincipalSansPontsTexte
        .Fields("TempsTotalPostesTexte").value = TempsTotalPostesSansPontsTexte
        .Fields("TempsTotalEgouttagesTexte").value = TempsTotalEgouttagesSansPontsTexte
        .Fields("TempsTotalGammeTexte").value = TempsTotalGammeSansPontsTexte
    
        '--- recherche le passage dans les bains principaux ---
        RecherchePassageBainsPrincipaux PassageAnodisation, _
                                                                   PassageSpectro, _
                                                                   PassageOr, _
                                                                   PassageNoir
        
        '--- affectation dans la base de donn�es ---
        .Fields("PassageAnodisation").value = IIf(PassageAnodisation = True, 1, 0)
        .Fields("PassageSpectro").value = IIf(PassageSpectro = True, 1, 0)
        .Fields("PassageOr").value = IIf(PassageOr = True, 1, 0)
        .Fields("PassageNoir").value = IIf(PassageNoir = True, 1, 0)
    
        '--- enregistrement de la partie redresseur ---
        .Fields("ModeUouI").value = ModeUouIEnCours
        For a = PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4
            .Fields("TempsPhase" & a).value = CTempsTexteEnSecondes(MEBTempsPhases(a).Text)
            If IsNumeric(TBTensionsPhases(a).Text) = True Then
                .Fields("UPhase" & a).value = CSng(TBTensionsPhases(a).Text)
            End If
            If IsNumeric(TBIntensitesPhases(a).Text) = True Then
                .Fields("IPhase" & a).value = CSng(TBIntensitesPhases(a).Text)
            End If
        Next a
    
    End With
    
    '--- valider l'enregistrement ---
    ADODCGammesAnodisation.Recordset.UpdateBatch adAffectAllChapters
    
    '--- actualisation ---
    CBActualiser_Click
    
    '--- curseur de la souris ---
    SourisEnAttente False
    
    '--- gestion des boutons ---
    GestionBoutons E_APRES_VALIDER
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
    
End Sub

Private Sub CBValider_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
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

Private Sub CBVerifierCoherenceGamme_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affichage du message gamme coh�rente ---
    If VerifierCoherenceGamme = True Then
        Bidon = AppelFenetre(F_MESSAGE, _
                                          TITRE_MESSAGES, _
                                          vbCrLf & vbCrLf & vbCrLf & "cs|LA GAMME EST COHERENTE" & vbCrLf & vbCrLf, _
                                          TYPES_MESSAGES.T_REMARQUE, _
                                          TYPES_BOUTONS.T_CONFIRMER, _
                                          EMPLACEMENT_FOCUS.E_SUR_CONFIRMER)
    End If

End Sub

Private Sub CBVerifierCoherenceGamme_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBVerifierCoherenceGamme_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CTOnglets_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    Select Case CTOnglets.CurrTab

        Case ONGLETS.O_RENSEIGNEMENTS
            '--- renseignements ---

        Case ONGLETS.O_GAMME_ANODISATION
            '--- gamme d'anodisation ---

        Case ONGLETS.O_CALCULS_PAR_APPRENTISSAGE
            '--- calculs par apprentissage ---
        
        Case Else

    End Select

End Sub

Private Sub DGMatieres_DblClick()
    On Error Resume Next
    InsertionMatiere
End Sub

Private Sub DGMatieres_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = vbDataErrContinue
End Sub

Private Sub DGMatieres_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- cadre de focus ---
    SFocusTableMatieres.Visible = True

    '--- affichage de la barre de s�lection ---
    With DGMatieres
        .CurrentCellVisible = True
        .Refresh
    End With

End Sub

Private Sub DGMatieres_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn
            InsertionMatiere
            KeyCode = 0: Shift = 0
        Case Else
    End Select
End Sub

Private Sub DGMatieres_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- cadre de focus ---
    SFocusTableMatieres.Visible = False

End Sub

Private Sub DGMatieres_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With DGMatieres 'pour fixer toujour la premi�re colonne
        .Col = 0
        .CurrentCellVisible = True
    End With
End Sub

Private Sub DGZones_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn
            InsertionDetail 0
            KeyCode = 0: Shift = 0
        Case Else
    End Select
End Sub

Private Sub DGZones_DblClick()
    On Error Resume Next
    InsertionDetail
End Sub

Private Sub DGZones_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = vbDataErrContinue
End Sub

Private Sub DGZones_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- cadre de focus ---
    SFocusTableZones.Visible = True

    '--- affichage de la barre de s�lection ---
    With DGZones
        .CurrentCellVisible = True
        .Refresh
    End With

End Sub

Private Sub DGZones_HeadClick(ByVal ColIndex As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- lancement de la requ�te ---
    With ADODCZones

        '--- requ�te ---
        Select Case DGZones.Columns(ColIndex)
            Case 2: .RecordSource = "SELECT Zones.* FROM Zones ORDER BY Numzone"
            Case Else: .RecordSource = "SELECT Zones.* FROM Zones ORDER BY NumZone"
        End Select

        '--- rafraichissement ---
        .Refresh
        .Recordset.MoveFirst

    End With

    '--- effacement de la s�lection de colonne ---
    DGZones.ClearSelCols

End Sub

Private Sub DGZones_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- cadre de focus ---
    SFocusTableZones.Visible = False

End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fen�tre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        Me.Refresh
        LectureDetailsGammesAnodisation
        If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
        PremiereActivation = True
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Marque ou restitue un enregistrement (fonction Bookmark)
' Entr�es : MarquageRestitution -> TRUE  = Marquage de l'enregistrement
'                                                       FALSE = Restitution de l'enregistrement
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MarqueEnregistrement(ByVal MarquageRestitution As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Static SignetEnregistrement As Variant

    With ADODCGammesAnodisation.Recordset
        If .EOF = False And .BOF = False Then
            If MarquageRestitution = True Then
                SignetEnregistrement = .Bookmark
            Else
                If IsEmpty(SignetEnregistrement) = False Then
                    If SignetEnregistrement > 0 Then
                        .Bookmark = SignetEnregistrement
                    End If
                End If
            End If
        End If
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affiche la grille de recherche
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheGrilleRecherche()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const HauteurPBCriteresRecherche As Integer = 1294
    
    '--- d�claration ---
    Dim HauteurGrilleRecherche As Long
    
    '--- affichage ---
    If RechercherSurGrille = False Then
        PBCriteresRecherche.Height = HauteurPBCriteresRecherche
        TDBGGrilleRecherche.Visible = False
        Me.Refresh
    Else
        PBCriteresRecherche.Height = PBBoutons.Top - PBCriteresRecherche.Top - MARGES.M_BORD_BAS - 10 * Screen.TwipsPerPixelY
        TDBGGrilleRecherche.Visible = True
    End If
    
    '--- hauteur de la grille de recherche ---
    HauteurGrilleRecherche = PBCriteresRecherche.Height - TDBGGrilleRecherche.Top - TDBGGrilleRecherche.Left - 5 * Screen.TwipsPerPixelY
    If HauteurGrilleRecherche > 0 Then
        'TDBGGrilleRecherche.Height = HauteurGrilleRecherche
    End If
    
    '--- placer le focus ---
    If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
    
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
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
        
        '--- affectation du mode intensit� ---
        ModeUouIEnCours = MODES_U_OU_I.M_INTENSITE
        
    End If

    '--- gestion des boutons ---
    If InterdireEvenements = False Then
        GestionBoutons E_MODIFICATION_EN_COURS
    End If

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

    '--- d�claration ---
    Dim a As Integer            'pour les boucles FOR...NEXT

    '--- affectation automatique du temps dans la gamme ---
    If InterdireEvenements = False Then
    
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
    
            With TDetailsGammesAnodisation(a)
                        
                '--- recherche de la zone d'anodisation ---
                If Trim(.Codezone) = "C13 � C16" Then
                    
                    '--- affectation dans le tableau ---
                    .TempsAuPosteTexte = "0" & LTempsTotalGammeRedresseur.Caption
                    .TempsAuPosteSecondes = CTempsTexteEnSecondes(LTempsTotalGammeRedresseur.Caption)
                    
                    '--- rafraichissement dans la grille ---
                    MSHFGDetailsGammesAnodisation.TextMatrix(a, COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE) = .TempsAuPosteTexte
                    
                    Exit For
                
                End If
                            
            End With
    
        Next a

    End If

End Sub

Private Sub MEBEditionDetailsGammesAnodisation_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim TexteComplet As String, _
            TexteSansMasque As String

    If InterdireEvenements = False Then
    
        '--- affectation ---
        With MEBEditionDetailsGammesAnodisation
            TexteComplet = .Text
            TexteSansMasque = .ClipText
        End With
    
        '--- analyse en fonction de chaque colonne ---
        With MSHFGDetailsGammesAnodisation
                    
            Select Case .Col
                
                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                    '--- temps au poste en texte ---
                    TDetailsGammesAnodisation(.Row).TempsAuPosteTexte = Replace(TexteComplet, "_", "0")
                    TDetailsGammesAnodisation(.Row).TempsAuPosteSecondes = CTempsTexteEnSecondes(TDetailsGammesAnodisation(.Row).TempsAuPosteTexte)
                    GestionBoutons E_MODIFICATION_EN_COURS
                
                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE
                    '--- temps d'alerte en texte ---
                    If Replace(TexteComplet, "_", "0") = "00:00:00" Then
                        TDetailsGammesAnodisation(.Row).TempsAlerteTexte = ""
                        TDetailsGammesAnodisation(.Row).TempsAlerteSecondes = 0
                    Else
                        TDetailsGammesAnodisation(.Row).TempsAlerteTexte = Replace(TexteComplet, "_", "0")
                        TDetailsGammesAnodisation(.Row).TempsAlerteSecondes = CTempsTexteEnSecondes(TDetailsGammesAnodisation(.Row).TempsAlerteTexte)
                    End If
                    GestionBoutons E_MODIFICATION_EN_COURS
                
                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                    '--- temps d'�gouttage en texte ---
                    TDetailsGammesAnodisation(.Row).TempsEgouttageTexte = Replace(TexteComplet, "_", "0")
                    TDetailsGammesAnodisation(.Row).TempsEgouttageSecondes = CTempsTexteEnSecondes(TDetailsGammesAnodisation(.Row).TempsEgouttageTexte)
                    GestionBoutons E_MODIFICATION_EN_COURS
                
                Case Else
    
            End Select
    
        End With

    End If

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
        .SelStart = 0          'met en surbrillance la s�lection saisie
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
    GestionBoutons E_MODIFICATION_EN_COURS
End Sub

Private Sub MEBTempsPhases_ValidationError(Index As Integer, InvalidText As String, StartPosition As Integer)
    On Error Resume Next
    MEBTempsPhases(Index).Text = Replace(InvalidText, "_", "0")
End Sub

Private Sub MSHFGDetailsGammesAnodisation_DblClick()
    On Error Resume Next
    InterdireEvenements = True
    EditionDetailsGammesAnodisation vbKeySpace  'simule un espace
    InterdireEvenements = False
End Sub

Private Sub MSHFGDetailsGammesAnodisation_GotFocus()
    On Error Resume Next
    SFocusTableDetailsGammesAnodisation.Visible = True
End Sub

Private Sub MSHFGDetailsGammesAnodisation_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete: EditionDetailsGammesAnodisation vbKeyBack  'simule un retour arri�re (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGDetailsGammesAnodisation_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionDetailsGammesAnodisation KeyAscii  'envoi de la touche frapp�e
End Sub

Private Sub MSHFGDetailsGammesAnodisation_LeaveCell()
    On Error Resume Next
    MEBEditionDetailsGammesAnodisation.Visible = False
End Sub

Private Sub MSHFGDetailsGammesAnodisation_LostFocus()
    On Error Resume Next
    SFocusTableDetailsGammesAnodisation.Visible = False
End Sub

Private Sub MSHFGDetailsGammesAnodisation_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- m�morisation de la ligne de d�part ---
    With MSHFGDetailsGammesAnodisation
        If Button = vbKeyLButton And .MouseCol = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NUM_LIGNES Then
            LigneDepartDeplacement = .MouseRow
        End If
    End With

End Sub

Private Sub MSHFGDetailsGammesAnodisation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- d�claration ---
    Dim TexteCellule As String
    Static MemTexteCellule As String
    
    '--- m�morisation de la ligne de d�part ---
    With MSHFGDetailsGammesAnodisation
        
        '--- RAZ des variables de d�placement ---
        If Button <> vbKeyLButton Then
            LigneDepartDeplacement = 0
            LigneArriveeDeplacement = 0
        End If
        
        '--- affectation ---
        TexteCellule = .TextMatrix(.MouseRow, .MouseCol)
        
        If TexteCellule <> MemTexteCellule Then
        
            '--- gestion de la bulle ---
            Select Case .MouseCol
            
                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                    '--- temps au poste en texte ---
                
                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE
                    '--- temps d'alerte en texte ---
                        
                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                    '--- temps d'�gouttage en texte ---
                
                Case Else
                    .ToolTipText = ""
        
            End Select
    
            '--- affectation ---
            MemTexteCellule = TexteCellule
    
        End If
    
    End With

End Sub

Private Sub MSHFGDetailsGammesAnodisation_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- m�morisation de la ligne d'arriv�e ---
    With MSHFGDetailsGammesAnodisation
        If Button = vbKeyLButton And .MouseCol = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NUM_LIGNES Then
            LigneArriveeDeplacement = .MouseRow
            If LigneDepartDeplacement > 0 And _
               LigneArriveeDeplacement > 0 And _
               LigneDepartDeplacement <> LigneArriveeDeplacement Then
                    DeplacementLigne
            End If
        End If
    End With

End Sub

Private Sub MSHFGDetailsGammesAnodisation_Scroll()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rendre invisible le champ d'�dition en cas de scrolling ---
    If MEBEditionDetailsGammesAnodisation.Visible = True Then
        MEBEditionDetailsGammesAnodisation.Visible = False
    End If

End Sub

Private Sub OBFormeGrille_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                                                      'pour les boucles FOR...NEXT
    Dim NbrFractionnements As Integer                       'nombre de fractionnement
    
    '--- changement de la forme d'affichage ---
    SourisEnAttente True
    With TDBGGrilleRecherche
        Select Case Index
            
            Case 0
                '--- remettre en pr�sentation normale ---
                .DataView = dbgNormalView               'pr�sentation normale
                '.Splits(0).AllowSizing = True               'autorise le fractionnement de la grille (petite rectangle noir en bas � gauche)
            
            Case 1
                '--- changement de la pr�sentation ---
                NbrFractionnements = .Splits.Count
                If NbrFractionnements > 1 Then
                    For a = 2 To NbrFractionnements
                        .Splits.Remove 1                         'effacer le fractionnement 1 quelque soit le nombre de fractionnements
                    Next a
                End If
                .DataView = dbgInvertedView              'pr�sentation invers�e
            
            Case Else
        End Select
    End With
    SourisEnAttente False

    '--- placer le focus sur la grille ---
    If TDBGGrilleRecherche.Visible = True Then TDBGGrilleRecherche.SetFocus

End Sub

Private Sub PBBoutons_Resize()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBValider.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBValider.Width
    CBAnnuler.Left = CBValider.Left - MARGES.M_ENTRE_BOUTONS - CBAnnuler.Width
    ADODCGammesAnodisation.Left = CBAnnuler.Left - MARGES.M_ENTRE_BOUTONS - ADODCGammesAnodisation.Width
    LRenseignements.Left = ADODCGammesAnodisation.Left
    CBNouveau.Left = ADODCGammesAnodisation.Left - MARGES.M_ENTRE_BOUTONS - CBNouveau.Width
    CBActualiser.Left = CBNouveau.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
    CBSupprimer.Left = CBActualiser.Left - MARGES.M_ENTRE_BOUTONS - CBSupprimer.Width
    CBVerifierCoherenceGamme.Left = CBSupprimer.Left - MARGES.M_ENTRE_BOUTONS - CBVerifierCoherenceGamme.Width
    AfficheGrilleRecherche
    
    '--- recalcul du focus apr�s d�placement ---
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
' R�le      : G�re l'�tats des boutons apr�s une action de l'op�rateur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionBoutons(ByVal Situation As ETATS_BOUTONS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case Situation
        
        Case ETATS_BOUTONS.E_CHARGEMENT_FENETRE
            '--- au chargement de la fenetre ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            ADODCGammesAnodisation.Enabled = True
            CBNouveau.Enabled = True
            CBActualiser.Enabled = True
            PBCriteresRecherche.Enabled = True
            CBCopieGammes.Enabled = True
            FNouveauNumGamme.Visible = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au d�chargement de la fen�tre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
            ADODCGammesAnodisation.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- apr�s valider ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBNouveau.Enabled = True
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
            CBCopieGammes.Enabled = True
            FNouveauNumGamme.Visible = True
        
        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
            ADODCGammesAnodisation.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- apr�s annuler ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBNouveau.Enabled = True
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
            CBCopieGammes.Enabled = True
            FNouveauNumGamme.Visible = True
        
        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            ADODCGammesAnodisation.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- apr�s actualiser ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBNouveau.Enabled = True
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
            CBCopieGammes.Enabled = True
            FNouveauNumGamme.Visible = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- apr�s modifier (� ne pas traiter si nouvel enregistrement) ---
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
            MarqueEnregistrement True
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBAnnuler.Enabled = True
            ADODCGammesAnodisation.Enabled = False
            CBNouveau.Enabled = False
            CBActualiser.Enabled = False
            CBSupprimer.Enabled = False
            PBCriteresRecherche.Enabled = False
            CBCopieGammes.Enabled = False
            FNouveauNumGamme.Visible = False
        
        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- apr�s nouveau ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            PBCriteresRecherche.Enabled = False
            CBCopieGammes.Enabled = False
            FNouveauNumGamme.Visible = False
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBAnnuler.Enabled = True
            ADODCGammesAnodisation.Enabled = False
            CBNouveau.Enabled = False
            CBActualiser.Enabled = False
            CBSupprimer.Enabled = False
            Me.TBNomGamme.SetFocus
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            ADODCGammesAnodisation.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- apr�s supprimer ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBNouveau.Enabled = True
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
            CBCopieGammes.Enabled = True
            FNouveauNumGamme.Visible = True
        
        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

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

Private Sub TBCommencantPar_GotFocus()
    On Error Resume Next
    With TBCommencantPar
        If .SelText = "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Sub TBContenant_GotFocus()
    On Error Resume Next
    With TBContenant
        If .SelText = "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With
End Sub

Private Sub TBDesignation_Change()
    On Error Resume Next
    With TBDesignation
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub TBCommencantPar_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri
            If RechercherSurGrille = True Then
                TDBGGrilleRecherche.SetFocus
            Else
                TBNumGamme.SetFocus
            End If
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBCommencantPar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case Succ(CBRechercherPar.ListIndex)
        Case IDX_RECHERCHER_PAR.IDX_NUM_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6                                'n� de gamme
        Case IDX_RECHERCHER_PAR.IDX_REF_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE, 30                                         'r�f�rence de la gamme
        Case IDX_RECHERCHER_PAR.IDX_DATE_CREATION_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_DATE_JJMMAAAA                'date de cr�ation
        Case IDX_RECHERCHER_PAR.IDX_NOM_GAMME: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE, 50                                        'nom de la gamme
        Case Else
    End Select
End Sub

Private Sub TBContenant_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri
            If RechercherSurGrille = True Then
                TDBGGrilleRecherche.SetFocus
            Else
                TBNumGamme.SetFocus
            End If
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBContenant_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES
End Sub

Private Sub MBDateGamme_Change()
    On Error Resume Next
    With MBDateGamme
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub MBDateGamme_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub MBDateGamme_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_DATE_JJMMAAAA
End Sub

Private Sub TBDesignation_GotFocus()
    On Error Resume Next
    Me.ActiveControl.MaxLength = ADODCGammesAnodisation.Recordset(Me.ActiveControl.DataField).DefinedSize
End Sub

Private Sub MEBEditionDetailsGammesAnodisation_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rendre visible ---
    SFocusTableDetailsGammesAnodisation.Visible = True

End Sub

Private Sub MEBEditionDetailsGammesAnodisation_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With MSHFGDetailsGammesAnodisation

        '--- analyse de la touche ---
        Select Case KeyCode

            Case vbKeyDown
                '--- fl�che basse ---
                .SetFocus
                If .Row < .Rows - 1 Then .Row = .Row + 1
                KeyCode = 0
            
            Case vbKeyUp
                '--- fl�che haute ---
                .SetFocus
                If .Row > .FixedRows Then .Row = .Row - 1
                KeyCode = 0
  
            Case Else
  
        End Select
  
    End With
  
End Sub

Private Sub MEBEditionDetailsGammesAnodisation_KeyPress(KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---

    With MSHFGDetailsGammesAnodisation

        '--- analyse de la touche ---
        Select Case KeyAscii

            Case vbKeyReturn
                '--- touche entr�e ---
                Select Case .Col

                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                        '--- temps au poste en texte ---
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                        
                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                        '--- temps d'�gouttage en texte ---
                        If .Row < .Rows - 1 Then .Row = .Row + 1
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                        
                    Case Else

                End Select

                '--- mettre le focus sur le tableau ---
                .SetFocus
                KeyAscii = 0

            Case Else
                Select Case .Col
                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 8
                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 5
                    Case Else
                End Select

        End Select

    End With

End Sub

Private Sub MEBEditionDetailsGammesAnodisation_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    
    '--- focus ---
    SFocusTableDetailsGammesAnodisation.Visible = False
    
    '--- rendre le contr�le texte invisible ---
    MEBEditionDetailsGammesAnodisation.Visible = False

    '--- construction de la grille ---
    GestionDetailsGammesAnodisation GG_AFFICHAGE
    
End Sub

Private Sub TBIntensitesPhases_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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
        .SelStart = 0          'met en surbrillance la s�lection saisie
        .SelLength = Len(.Text)
    End With
    
    '--- gestion des boutons ---
    If InterdireEvenements = False Then
        GestionBoutons E_MODIFICATION_EN_COURS
    End If

End Sub

Private Sub TBIntensitesPhases_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBIntensitesPhases_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 4
    GestionBoutons E_MODIFICATION_EN_COURS
End Sub

Private Sub TBIntensitesPhases_LostFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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

Private Sub TBMatieres_Change(Index As Integer)
    On Error Resume Next
    With TBMatieres(Index)
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub TBMatieres_GotFocus(Index As Integer)
    On Error Resume Next
    Me.ActiveControl.MaxLength = ADODCGammesAnodisation.Recordset(Me.ActiveControl.DataField).DefinedSize
End Sub

Private Sub TBMatieres_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBMatieres_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE
End Sub

Private Sub TBNomGamme_Change()
    On Error Resume Next
    With TBNomGamme
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub TBNomGamme_GotFocus()
    On Error Resume Next
    Me.ActiveControl.MaxLength = ADODCGammesAnodisation.Recordset(Me.ActiveControl.DataField).DefinedSize
End Sub

Private Sub TBNomGamme_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBNomGamme_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Gestion des d�tails des gammes d'anodisation
' Entr�es : EtatSouhaite -> Fonction de l'�num�ration GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsGammesAnodisation(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---

    '--- d�claration ---
    Dim TypeCouleur As Boolean
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer, _
            PtrLigne As Integer, _
            NumZoneDepart As Integer, _
            NumZoneArrivee As Integer, _
            NumPont As Integer
    Dim FicheVide As ImgDetailsGammesProduction, _
            TCopieDetailsgammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesProduction

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des d�tails ---
            Erase TDetailsGammesAnodisation()

            '--- initialisation de la grille des d�tails ---
            With MSHFGDetailsGammesAnodisation

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_GAMMES_PRODUCTION + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     '�paisseur de lignes modifi�es ligne par ligne
                .RowHeight(0) = 750                                        '�paisseur des titres
                .RowHeightMin = 315
                .Row = 0

                '--- param�trages de chaque colonne ---
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Code de la zone"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_LIBELLE_ZONE
                .ColWidth(.Col) = 37.78 * EPAISSEUR_CARACTERE: .Text = "Libell� de la zone"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Temps au poste"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Temps d'alerte"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Temps d'�gouttage"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_PONT
                .ColWidth(.Col) = 5 * EPAISSEUR_CARACTERE: .Text = "Pont"
                .ColAlignment(.Col) = flexAlignCenterCenter

                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N� de lignes, vidage des champs ---
                For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                
                    '--- N� de lignes ---
                    .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NUM_LIGNES
                    .RowHeight(a) = 315                    '�paisseur des lignes
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_GAMMES_PRODUCTION.C_PONT
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.VERT_1, COULEURS.CYAN_1)
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE

                .Redraw = True

            End With
                
        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                TDetailsGammesAnodisation(a) = FicheVide
            Next a
            With MSHFGDetailsGammesAnodisation
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- zone des donn�es dans le tableau ---
            PtrLigne = 1
            For a = 1 To UBound(TTempEnrDetailsGammesAnodisation())
                With TTempEnrDetailsGammesAnodisation(a)
                    TDetailsGammesAnodisation(PtrLigne).NumZone = .NumZone
                    If .NumZone > 0 Then
                        
                        TDetailsGammesAnodisation(PtrLigne).Codezone = TZones(.NumZone).Codezone
                        TDetailsGammesAnodisation(PtrLigne).LibelleZone = TZones(.NumZone).LibelleZone
                        TDetailsGammesAnodisation(PtrLigne).TempsAuPosteTexte = .TempsAuPosteTexte
                        TDetailsGammesAnodisation(PtrLigne).TempsAlerteTexte = .TempsAlerteTexte
                        TDetailsGammesAnodisation(PtrLigne).TempsEgouttageTexte = .TempsEgouttageTexte
                        TDetailsGammesAnodisation(PtrLigne).TempsAuPosteSecondes = .TempsAuPosteSecondes
                        TDetailsGammesAnodisation(PtrLigne).TempsAlerteSecondes = .TempsAlerteSecondes
                        TDetailsGammesAnodisation(PtrLigne).TempsEgouttageSecondes = .TempsEgouttageSecondes
                        
                        Inc PtrLigne
                    
                    End If
                End With
            Next a

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des donn�es ---
            PtrLigne = 1
            For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                If TDetailsGammesAnodisation(a).NumZone <> 0 Then
                    TCopieDetailsgammesAnodisation(PtrLigne) = TDetailsGammesAnodisation(a)
                    Inc PtrLigne
                End If
            Next a
            For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                TDetailsGammesAnodisation(a) = TCopieDetailsgammesAnodisation(a)
            Next a

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            
            '--- ne pas afficher la partie redresseur par d�faut ---
            FRedresseurs.Visible = False
            
            With MSHFGDetailsGammesAnodisation

                '--- m�morisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                    .Row = a
                    If TDetailsGammesAnodisation(a).NumZone = 0 Then
                        TDetailsGammesAnodisation(a) = FicheVide
                        For b = 1 To NBR_COLONNES_DETAILS_GAMMES_PRODUCTION
                            .Col = b
                            .Text = ""
                        Next b
                    Else
                        
                        '--- affichage de la partie redresseur ---
                        If Trim(TDetailsGammesAnodisation(a).Codezone) = "C13 � C16" Then
                            FRedresseurs.Visible = True
                        End If
                        
                        '--- affichage ---
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                        .Text = TDetailsGammesAnodisation(a).Codezone
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_LIBELLE_ZONE
                        .Text = TDetailsGammesAnodisation(a).LibelleZone
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                        .Text = TDetailsGammesAnodisation(a).TempsAuPosteTexte
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE
                        If TDetailsGammesAnodisation(a).TempsAlerteTexte = "00:00:00" Then TDetailsGammesAnodisation(a).TempsAlerteTexte = ""
                        .Text = TDetailsGammesAnodisation(a).TempsAlerteTexte
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                        .Text = TDetailsGammesAnodisation(a).TempsEgouttageTexte
                        
                        '--- affectation des num�ros de zones pour l'affichage du pont ---
                        NumZoneDepart = TDetailsGammesAnodisation(a).NumZone
                        If a = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                            NumZoneArrivee = 0
                        Else
                            If TDetailsGammesAnodisation(a + 1).Codezone = "" Then
                                NumZoneArrivee = 0
                            Else
                                NumZoneArrivee = TDetailsGammesAnodisation(a + 1).NumZone
                            End If
                        End If
                        
                        '--- affichage du pont ---
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_PONT
                        If NumZoneDepart > 0 And NumZoneArrivee > 0 Then
                            NumPont = RechercheNumPontChoisiDansPremisse(NumZoneDepart, NumZoneArrivee)
                            If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                                .Text = "P" & NumPont
                                .CellBackColor = IIf(TypeCouleur = False, COULEURS.VERT_1, COULEURS.CYAN_1)
                            Else
                                .Text = "*"
                                .CellBackColor = COULEURS.ROUGE_1
                            End If
                        Else
                            .Text = ""
                        End If
                    
                    End If
                
                    '--- affectation de la variable permettant l'alternance de couleurs sur le tableau ---
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- restitution des valeurs ligne, colonne ---
                .Redraw = True
                .Row = MemLigne
                .Col = MemColonne
                .FocusRect = flexFocusHeavy

            End With

            '--- affichage des temps de la gamme ---
            AffichageTempsGamme
            AffichageCalculsParApprentissage
        
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Insertion d'un d�tail dans la grille des d�tails
' Entr�es : Codezone -> Code du zone
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InsertionDetail(Optional ByVal Codezone As Variant)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    Dim ChaineDeRecherche As String
    Dim FicheVide As ImgDetailsGammesProduction

    '--- lancer la modification ---
    GestionBoutons E_MODIFICATION_EN_COURS

    If IsMissing(Codezone) = True Then

        '--- les donn�es viennent de la grille des codes ---
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            With TDetailsGammesAnodisation(a)
                If .Codezone = "" Then
                    
                    '--- affectation ---
                    .NumZone = ADODCZones.Recordset("NumZone").value
                    .Codezone = TZones(.NumZone).Codezone
                    .LibelleZone = TZones(.NumZone).LibelleZone
                    
                    If TEtatsPostes(TZones(.NumZone).NumPremierPoste).DefinitionPoste.AvecTemps = True Then
                        .TempsAuPosteTexte = "00:00:00"
                    Else
                        .TempsAuPosteTexte = PAS_DE_TEMPS
                    End If
                    
                    If TEtatsPostes(TZones(.NumZone).NumPremierPoste).DefinitionPoste.AvecTemps = True Then
                        .TempsAlerteTexte = ""
                    Else
                        .TempsAlerteTexte = PAS_DE_TEMPS
                    End If
                    
                    If TEtatsPostes(TZones(.NumZone).NumPremierPoste).DefinitionPoste.AvecEgouttage = True Then
                        .TempsEgouttageTexte = "00:00"
                    Else
                        .TempsEgouttageTexte = PAS_DE_TEMPS
                    End If
                    
                    .TempsAuPosteSecondes = 0
                    .TempsAlerteSecondes = 0
                    .TempsEgouttageSecondes = 0
                    
                    With MSHFGDetailsGammesAnodisation
                        If .RowIsVisible(a) = False Then
                            .TopRow = a
                        End If
                        .Row = a
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                    End With

                    Exit For

                End If
            End With
        Next a
        GestionDetailsGammesAnodisation GG_AFFICHAGE
        MSHFGDetailsGammesAnodisation.SetFocus

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : D�place une ligne dans la grille des d�tails
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DeplacementLigne()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
            PtrLigne As Integer
    Dim TFicheDepart As ImgDetailsGammesProduction, _
            TCopieDetailsgammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesProduction

    If LigneDepartDeplacement > 0 And LigneDepartDeplacement < NBR_LIGNES_DETAILS_GAMMES_PRODUCTION And _
       LigneArriveeDeplacement > 0 And LigneArriveeDeplacement < NBR_LIGNES_DETAILS_GAMMES_PRODUCTION And _
       LigneDepartDeplacement <> LigneArriveeDeplacement Then

        '--- affectation ---
        TFicheDepart = TDetailsGammesAnodisation(LigneDepartDeplacement)

        '--- suppression � la ligne de d�part ---
        PtrLigne = 1
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            If a <> LigneDepartDeplacement Then
                TCopieDetailsgammesAnodisation(PtrLigne) = TDetailsGammesAnodisation(a)
                Inc PtrLigne
            End If
        Next a

        '--- zone dans le tableau ---
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            TDetailsGammesAnodisation(a) = TCopieDetailsgammesAnodisation(a)
        Next a

        '--- fixation de l'arriv�e en fonction du sens de d�placement ---
        If LigneArriveeDeplacement > LigneDepartDeplacement Then
            LigneArriveeDeplacement = Pred(LigneArriveeDeplacement)
        End If
        If LigneArriveeDeplacement < 1 Then LigneArriveeDeplacement = 1
        If LigneArriveeDeplacement > NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then LigneArriveeDeplacement = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION

        '--- insertion � la ligne d'arriv�e ---
        PtrLigne = 1
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            If a = LigneArriveeDeplacement Then
                TCopieDetailsgammesAnodisation(PtrLigne) = TFicheDepart
                Inc PtrLigne
            End If
            If PtrLigne <= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                TCopieDetailsgammesAnodisation(PtrLigne) = TDetailsGammesAnodisation(a)
                Inc PtrLigne
            End If
            If PtrLigne >= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then Exit For
        Next a

        '--- zone dans le tableau ---
        For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            TDetailsGammesAnodisation(a) = TCopieDetailsgammesAnodisation(a)
        Next a

        '--- reconstruction de la grille ---
        GestionDetailsGammesAnodisation GG_AFFICHAGE

        '--- gestion des boutons ---
        GestionBoutons E_MODIFICATION_EN_COURS
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Lecture des d�tails des gammes d'anodisation
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureDetailsGammesAnodisation()

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- d�claration ---
    Dim a As Integer
    Dim NumGamme As String

    If MemDernierBouton <> ETATS_BOUTONS.E_AVANT_NOUVEAU And _
       MemDernierBouton <> ETATS_BOUTONS.E_APRES_NOUVEAU Then

        '--- curseur de la souris ---
        SourisEnAttente True

        '--- vidage des grilles ---
        GestionDetailsGammesAnodisation GG_VIDAGE
        GestionDetailsGammesAnodisation GG_AFFICHAGE

        With ADODCGammesAnodisation.Recordset

            If Not .BOF And Not .EOF Then

                If .status = adRecOK Then

                    If IsError(.Fields("NumGamme")) = False Then

                        '--- affectation ---
                        NumGamme = .Fields("NumGamme")

                        '--- recherche des d�tails des gammes d'anodisation ---
                        If RechercheDetailsGammesAnodisation(NumGamme) = TROUVE Then
                            GestionDetailsGammesAnodisation GG_TRANSFERT_DONNEES
                            GestionDetailsGammesAnodisation GG_AFFICHAGE
                        End If

                    End If

                End If

            End If

        End With

        '--- curseur de la souris ---
        SourisEnAttente False

    End If

    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:

    '--- curseur de la souris ---
    SourisEnAttente False

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Enregistrement des d�tails des gammes d'anodisation
' Entr�es : NumGamme -> Num�ro de la gamme
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistrementDetailsGammesAnodisation(ByVal NumGamme As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    '--- compression / construction ---
    GestionDetailsGammesAnodisation GG_COMPRESSION
    GestionDetailsGammesAnodisation GG_AFFICHAGE

    If NumGamme <> "" Then
        
        '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With

        '--- lancement de la requ�te ---
        With Enregistrement
            .CursorLocation = adUseServer
            .MaxRecords = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
            Requete = "SELECT DetailsGammesAnodisation.* FROM DetailsGammesAnodisation WHERE (NumGamme = '" & NumGamme & "')"
            .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        End With

        With Enregistrement

            '--- enregistrement des nouveaux d�tails ---
            For a = 1 To UBound(TDetailsGammesAnodisation())
                If TDetailsGammesAnodisation(a).NumZone <> 0 Then
                    
                    .AddNew
                    
                    !NumGamme = NumGamme
                    !NumLigne = a
                    !NumZone = TDetailsGammesAnodisation(a).NumZone
                    
                    !TempsAuPosteTexte = TDetailsGammesAnodisation(a).TempsAuPosteTexte
                    !TempsAlerteTexte = TDetailsGammesAnodisation(a).TempsAlerteTexte
                    !TempsEgouttageTexte = TDetailsGammesAnodisation(a).TempsEgouttageTexte
                    
                    !TempsAuPosteSecondes = TDetailsGammesAnodisation(a).TempsAuPosteSecondes
                    !TempsAlerteSecondes = TDetailsGammesAnodisation(a).TempsAlerteSecondes
                    !TempsEgouttageSecondes = TDetailsGammesAnodisation(a).TempsEgouttageSecondes
                    
                    .Update
                
                Else
                    Exit For
                End If
            Next a

        End With

        '--- fermeture de l'enregistrement / connexion ---
        Enregistrement.Close
        ConnexionBDAnodisationSQL.Close

    End If

    '--- affectation ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Suppression des d�tails des gammes d'anodisation
' Entr�es : NumGamme -> Num�ro de la gamme
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SuppressionDetailsGammesAnodisation(ByVal NumGamme As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    If NumGamme <> "" Then
        
        '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
        
        '--- lancement de la requ�te ---
        With Enregistrement
            .CursorLocation = adUseServer
            Requete = "DELETE FROM DetailsGammesAnodisation WHERE (NumGamme = '" & NumGamme & "')"
            .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        End With
    
        '--- fermeture de la connexion ---
        ConnexionBDAnodisationSQL.Close
    
    End If

    '--- affectation ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Suppression d'une gammee d'anodisation
' Entr�es : NumGamme -> Num�ro de la gamme
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SuppressionGammesAnodisation(ByVal NumGamme As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset

    If NumGamme <> "" Then
        
        '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
        
        '--- lancement de la requ�te ---
        With Enregistrement
            .CursorLocation = adUseServer
            Requete = "DELETE FROM GammesAnodisation WHERE (NumGamme = '" & NumGamme & "')"
            .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        End With
    
        '--- fermeture de la connexion ---
        ConnexionBDAnodisationSQL.Close
    
    End If

    '--- affectation ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : V�rifie la coh�rence de la gamme
' Entr�es :
' Retours : VerifierCoherenceGamme -> FALSE = La gamme n'est pas coh�rente
'                                                                TRUE = La gamme est coh�rente
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function VerifierCoherenceGamme() As Boolean

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer, _
           b As Integer, _
           c As Integer, _
           NumDerniereLigneGamme As Integer, _
           NumZoneDepart As Integer, _
           NumZoneArrivee As Integer, _
           NumPosteDepart As Integer, _
           NumPosteArrivee As Integer
    Dim CodeZonePremiereLigneGamme As String, _
           CodeZoneDerniereLigneGamme As String
    
    '--- forcer la valeur de retour comme si la gamme �tait coh�rente ---
    VerifierCoherenceGamme = True
    
    '--- compacter la gamme ---
    GestionDetailsGammesAnodisation GG_COMPRESSION
    GestionDetailsGammesAnodisation GG_AFFICHAGE
    
    '--- code de la zone de la premi�re ligne ---
    CodeZonePremiereLigneGamme = TDetailsGammesAnodisation(1).Codezone
    If CodeZonePremiereLigneGamme <> TZones(1).Codezone Then
        If AppelFenetre(F_MESSAGE, _
                                TITRE_MESSAGES, _
                                vbCrLf & "cs|INCOHERENCE DE LA GAMME" & vbCrLf & vbCrLf & _
                                "La premi�re zone de la gamme doit �tre" & vbCrLf & _
                                "imp�rativement la zone des postes de CHARGEMENT" & vbCrLf & vbCrLf & _
                                "cs|Cette gamme ne pourra pas �tre lanc�e", _
                                TYPES_MESSAGES.T_AVERTISSEMENT, _
                                TYPES_BOUTONS.T_CONFIRMER, _
                                EMPLACEMENT_FOCUS.E_SUR_CONFIRMER) = vbOK Then
        
            '--- gamme non coh�rente ---
            VerifierCoherenceGamme = False
        
        End If
    End If
    
    '--- v�rification de la premi�re ligne de la gamme ---
    For a = LBound(TDetailsGammesAnodisation()) To UBound(TDetailsGammesAnodisation())
        With TDetailsGammesAnodisation(a)
    
            '--- analyse des pr�misses ---
            If a <> UBound(TDetailsGammesAnodisation()) Then
                If TDetailsGammesAnodisation(a + 1).NumZone > 0 Then
                
                    '--- affectation ---
                    NumZoneDepart = TDetailsGammesAnodisation(a).NumZone
                    NumZoneArrivee = TDetailsGammesAnodisation(a + 1).NumZone

                    '--- v�rification de l'existence des pr�misses ---
                    For b = TZones(NumZoneDepart).NumPremierPoste To TZones(NumZoneDepart).NumDernierPoste
                        For c = TZones(NumZoneArrivee).NumPremierPoste To TZones(NumZoneArrivee).NumDernierPoste
            
                            '--- affectation ---
                            NumPosteDepart = b
                            NumPosteArrivee = c
            
                            '--- contr�le ---
                            If TPremisses(NumPosteDepart, NumPosteArrivee).PremisseCodee = "" Then
                                If AppelFenetre(F_MESSAGE, _
                                                        TITRE_MESSAGES, _
                                                        "cs|INCOHERENCE DE LA GAMME" & vbCrLf & vbCrLf & _
                                                        "La r�gle permettant le transfert " & vbCrLf & _
                                                        "du poste de d�part = " & TEtatsPostes(b).DefinitionPoste.NomPoste & " - " & TEtatsPostes(b).DefinitionPoste.LibellePoste & vbCrLf & _
                                                        "au poste d'arriv�e  = " & TEtatsPostes(c).DefinitionPoste.NomPoste & " - " & TEtatsPostes(c).DefinitionPoste.LibellePoste & vbCrLf & _
                                                        "cs|N'EXISTE PAS. CETTE REGLE EST NECESSAIRE" & vbCrLf & vbCrLf & _
                                                        "cs|Il faut g�n�rer cette r�gle dans l'�cran des pr�misses", _
                                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                                        TYPES_BOUTONS.T_CONFIRMER, _
                                                        EMPLACEMENT_FOCUS.E_SUR_CONFIRMER) = vbOK Then
            
                                    '--- gamme non coh�rente ---
                                    VerifierCoherenceGamme = False
                                
                                End If
                            End If
            
                        Next c
                    Next b
                
                End If
            End If

            '--- recherche du num�ro de la derni�re ligne ---
            If .NumZone > 0 Then
                NumDerniereLigneGamme = a
            Else
                Exit For
            End If
    
        End With
    Next a
    
    '--- code de la zone de la derni�re ligne ---
    If NumDerniereLigneGamme > 0 Then
        CodeZoneDerniereLigneGamme = TDetailsGammesAnodisation(NumDerniereLigneGamme).Codezone
        
        If CodeZoneDerniereLigneGamme <> TZones(1).Codezone And CodeZoneDerniereLigneGamme <> "D1 � D2" Then
            If AppelFenetre(F_MESSAGE, _
                                    TITRE_MESSAGES, _
                                    vbCrLf & "cs|INCOHERENCE DE LA GAMME" & vbCrLf & vbCrLf & _
                                    "La derni�re zone de la gamme doit �tre imp�rativement" & vbCrLf & _
                                    "la zone de CHARGEMENT ou DECHARGEMENT" & vbCrLf & vbCrLf & _
                                    "cs|Cette gamme ne pourra pas �tre lanc�e", _
                                    TYPES_MESSAGES.T_AVERTISSEMENT, _
                                    TYPES_BOUTONS.T_CONFIRMER, _
                                    EMPLACEMENT_FOCUS.E_SUR_CONFIRMER) = vbOK Then
            
                                        '--- gamme non coh�rente ---
                                        VerifierCoherenceGamme = False
            
            End If
        End If
    End If

    '--- contr�le du nombre de lignes de la gamme ---
    If NumDerniereLigneGamme < 3 Then
        If AppelFenetre(F_MESSAGE, _
                                TITRE_MESSAGES, _
                                "cs|INCOHERENCE DE LA GAMME" & vbCrLf & vbCrLf & _
                                "Votre gamme ne comporte pas assez de lignes" & vbCrLf & _
                                "pour �tre exploiter correctement (3 lignes minimum)" & vbCrLf & vbCrLf & _
                                "cs|ATTENTION AUX RISQUES DE COLLISION" & vbCrLf & vbCrLf & _
                                "cs|Cette gamme ne pourra pas �tre lanc�e", _
                                TYPES_MESSAGES.T_AVERTISSEMENT, _
                                TYPES_BOUTONS.T_CONFIRMER, _
                                EMPLACEMENT_FOCUS.E_SUR_CONFIRMER) = vbOK Then
                                        
                                    '--- gamme non coh�rente ---
                                    VerifierCoherenceGamme = False
        
        End If
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le prochain num�ro de gamme
' Entr�es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ProchainNumGamme() As String

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- ouverture de la connexion � la base de donn�es de l'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- recherche du dernier num�ro ---
    With Enregistrement

        '--- ouverture / pointer le premier enregistrement ---
        .CursorLocation = adUseServer
        .MaxRecords = 1
        Requete = "SELECT MAX(NumGamme) AS Expr1 FROM gammesAnodisation"
        .Open Requete, ConnexionBDAnodisationSQL, adOpenStatic, adLockOptimistic, adCmdText
        .MoveFirst

        '--- affectation ---
        ProchainNumGamme = Right(String(6, "0") & CStr(CLng(Trim(Enregistrement("Expr1"))) + 1), 6)

    End With
    
    '--- fermeture de l'enregistrement / connexion ---
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Change le curseur de la souris en fonction de l'attente
' Entr�es : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par d�faut
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub SourisEnAttente(ByVal AttenteOuiNon As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- changement du curseur ---
    If AttenteOuiNon = True Then
        Me.MousePointer = vbHourglass
        CTOnglets.MousePointer = vbHourglass
    Else
        Me.MousePointer = vbDefault
        CTOnglets.MousePointer = vbDefault
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue le param�trage de la fen�tre
' Entr�es :                   TravailSurGrille -> FALSE = Travail sur la fiche
'                                                                  TRUE  = Travail sur la grille de recherche
'                                     RechercherPar -> Valeur du champ TBRechercherPar
'                                  CommencantPar -> Valeur du champ TBCommencantPar
'                                            Contenant -> Valeur du champ TBContenant
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal TravailSurGrille As Boolean, _
                                                   ByVal RechercherPar As Integer, _
                                                   ByVal CommencantPar As String, _
                                                   ByVal Contenant As String)
                                               
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- recherche sur grille ---
    RechercherSurGrille = False
    If TravailSurGrille = True Then
        CBRechercherSurGrille_Click
    End If
    
    '--- rechercher par ---
    If RechercherPar > 0 Then
        CBRechercherPar.ListIndex = Pred(RechercherPar)
    Else
        CBRechercherPar.Text = CBRechercherPar.List(0)
    End If
    
    '--- commen�ant par ---
    TBCommencantPar.Text = CommencantPar
    
    '--- contenant ---
    TBContenant.Text = Contenant
    
    '--- initialisation des champs / grilles ---
    GestionGrilleRecherche GG_INITIALISATION
    GestionGrilleRecherche GG_AFFICHAGE
    
    '--- Initialise les champs de la partie redresseur ---
    InitialisationChampsRedresseur
    
    '--- lancement de la recherche ---
    LanceRechercheOuTri
    
    '--- lancement du timer ---
    TimerSimulationEntreeCharge.Enabled = True

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Initialise la fen�tre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- d�claration ---

    '--- affectation ---
  
    '--- divers sur la fen�tre ---
    With Me
        .Picture = ImgFondBleu1
        .WindowState = vbMaximized
    End With
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- divers sur ADO ---

    '--- divers sur les renseignements ---
    LRenseignements.BackColor = COULEURS.CYAN_0

    '--- divers sur la grille des articles des gammes d'anodisation ---
    With DGZones
        .BackColor = COULEURS.JAUNE_0
        .ForeColor = COULEURS.BLEU_5
    End With

    '--- gestion des d�tails ---
    GestionDetailsGammesAnodisation GG_INITIALISATION

    '--- affichage des temps de gamme ---
    AffichageTempsGamme
    AffichageCalculsParApprentissage
    
    '--- affectation ---
    CTOnglets.CurrTab = ONGLETS.O_RENSEIGNEMENTS

    '--- gestion de l'�tats des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE

    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:

    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : D�charge la fen�tre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False

    '--- curseur souris par d�faut ---
    SourisEnAttente False
    
    '--- neutralisation du timer ---
    With TimerSimulationEntreeCharge
        .Enabled = False
        .Interval = 0
    End With

    '--- d�chargement de la fen�tre ---
    Me.Visible = False
    Unload Me
    Set OccFGammesAnodisation = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Lance une recherche en fonction des crit�res
' Entr�es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LanceRechercheOuTri()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
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
    CommencantPar = TBCommencantPar.Text
    Contenant = TBContenant.Text
    IdxRecherchePar = Succ(CBRechercherPar.ListIndex)
    If IdxRecherchePar < 1 Then IdxRecherchePar = 1
    RechercherPar = Choose(IdxRecherchePar, _
                                              "NumGamme", _
                                              "RefGamme", _
                                              "DateCreationGamme", _
                                              "NomGamme")
            
    '--- d�but de la requ�te ---
    RequeteSQL = "SELECT GammesAnodisation.* FROM GammesAnodisation "

    '--- modification pour le cas du num�ro de la gamme d'anodisation ---
    Select Case IdxRecherchePar
        Case IDX_RECHERCHER_PAR.IDX_NUM_GAMME
            '--- cas du num�ro de la gamme d'anodisation ---
            If CommencantPar <> "" Then
                CommencantPar = Right(FORMAT_NUM_GAMME_ANODISATION & CommencantPar, 6)
            End If
        Case Else
    End Select
    
    If IdxRecherchePar = IDX_RECHERCHER_PAR.IDX_DATE_CREATION_GAMME Then
        
        '--- filtres pour la date ---
        Filtre1 = "(CONVERT(VARCHAR(10), " & RechercherPar & ", 103) LIKE '" & CommencantPar & "%') "
        Filtre2 = "(CONVERT(VARCHAR(10), " & RechercherPar & ", 103) LIKE '%" & Contenant & "%') "
    
    Else
        
        '--- filtres pour chaines de caract�res ---
        Filtre1 = "(" & RechercherPar & " LIKE '" & CommencantPar & "%') "
        Filtre2 = "(" & RechercherPar & " LIKE '%" & Contenant & "%') "
    
    End If

    '--- construction du filtre ---
    If CommencantPar = "" And Contenant = "" Then
    ElseIf CommencantPar <> "" And Contenant = "" Then
        RequeteSQL = RequeteSQL & "WHERE " & Filtre1
    ElseIf CommencantPar = "" And Contenant <> "" Then
        RequeteSQL = RequeteSQL & "WHERE " & Filtre2
    Else
        RequeteSQL = RequeteSQL & "WHERE " & Filtre1 & " AND " & Filtre2
    End If
    
    '--- fin de la requ�te ---
    RequeteSQL = RequeteSQL & "ORDER BY " & RechercherPar
    Select Case IdxRecherchePar
        Case 1: RequeteSQL = RequeteSQL & ", DateCreationGamme DESC"                          'NumGamme
        Case 2: RequeteSQL = RequeteSQL & ", NumGamme"                                                  'RefGamme
        Case 3: RequeteSQL = RequeteSQL & ", NumGamme"                                                  'DateCreationGamme
        Case 4: RequeteSQL = RequeteSQL & ", NumGamme, DateCreationGamme DESC"    'NomGamme
        Case Else
    End Select

    'Debug.Print RequeteSQL
    With ADODCGammesAnodisation
        
        '--- application de la requ�te ---
        .Recordset.Cancel
        If .RecordSource <> RequeteSQL Then
            .RecordSource = RequeteSQL
            .Refresh
        Else
            .Recordset.Requery
        End If
        
        '--- message si fiche non trouv�e ---
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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet l'�dition des d�tails des gammes d'anodisation
' Entr�es : KeyAscii -> Code ASCII de la touche frapp�e
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionDetailsGammesAnodisation(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---

    '--- pas d'�dition des champs si la ligne est vide ---
    If TDetailsGammesAnodisation(MSHFGDetailsGammesAnodisation.Row).NumZone = 0 Then
        Exit Sub
    End If

    '--- pas d'�dition des champs si pas de temps au postes ou pas d'�gouttage ---
    With MSHFGDetailsGammesAnodisation
        Select Case .Col
            
            Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                '--- temps au poste en texte ---
                If TEtatsPostes(TZones(TDetailsGammesAnodisation(.Row).NumZone).NumPremierPoste).DefinitionPoste.AvecTemps = False Then
                    Exit Sub
                End If
            
            Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE
                '--- temps d'alerte en texte ---
                If TEtatsPostes(TZones(TDetailsGammesAnodisation(.Row).NumZone).NumPremierPoste).DefinitionPoste.AvecTemps = False Then
                    Exit Sub
                End If
            
            Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                '--- temps d'�gouttage entexte ---
                If TEtatsPostes(TZones(TDetailsGammesAnodisation(.Row).NumZone).NumPremierPoste).DefinitionPoste.AvecEgouttage = False Then
                    Exit Sub
                End If
            
            Case Else
        End Select
    End With

    '--- �dition uniquement sur les bonnes colonnes ---
    Select Case MSHFGDetailsGammesAnodisation.Col

        Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE, _
                 COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE, _
                 COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE

            With MEBEditionDetailsGammesAnodisation

                '--- affiche le contr�le texte au bon endroit (dans la cellule) ---
                .Move MSHFGDetailsGammesAnodisation.Left + MSHFGDetailsGammesAnodisation.CellLeft, _
                           MSHFGDetailsGammesAnodisation.Top + MSHFGDetailsGammesAnodisation.CellTop, _
                           MSHFGDetailsGammesAnodisation.CellWidth, _
                           MSHFGDetailsGammesAnodisation.CellHeight

                '--- param�tres de contr�le texte en fonction de la cellule ---
                .Mask = ""
                .Text = ""
                Select Case MSHFGDetailsGammesAnodisation.Col
                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE: .Mask = "##:##:##"
                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE: .Mask = "##:##:##"
                    Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE: .Mask = "##:##"
                    Case Else
                End Select

                '--- analyse du caract�re qui a �t� tap� ---
                Select Case KeyAscii

                    Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                        '--- du code 0 � l'espace (sauf retour arri�re, Entr�e) cela signifie une modification du texte en cours ---
                        .SelText = MSHFGDetailsGammesAnodisation.Text
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .Visible = True
                        .SetFocus

                    Case vbKeyBack
                        '--- touche retour arri�re ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        DoEvents
                        MEBEditionDetailsGammesAnodisation_Change

                    Case vbKeyReturn
                        '--- touche Entr�e ---
                        With MSHFGDetailsGammesAnodisation
                            Select Case .Col
                                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE: .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE
                                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_ALERTE_TEXTE: .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                                Case COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                                    If .Row < .Rows - 1 Then .Row = .Row + 1
                                    .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                                Case Else
                            End Select
                        End With

                    Case Else
                        '--- tout autre �l�ment signifie le remplacement du texte en cours ---
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
' R�le      : Affichage des calculs par apprentissage
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageCalculsParApprentissage()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer
    Dim TGammesAnodisation As EnrGammesAnodisation
    
    '--- initialisation des variables ---
    TempsMouvementsAvantPostePrincipalSecondes = 0
    TempsAvantPostePrincipalAvecPontsSecondes = 0
    TempsPostePrincipalAvecPontsSecondes = 0
    TempsMouvementsApresPostePrincipalSecondes = 0
    TempsApresPostePrincipalAvecPontsSecondes = 0
    TempsTotalPostesAvecPontsSecondes = 0
    TempsTotalEgouttagesAvecPontsSecondes = 0
    TempsTotalMouvementsSecondes = 0
    TempsTotalGammeAvecPontsSecondes = 0
    
    TempsMouvementsAvantPostePrincipalTexte = ""
    TempsAvantPostePrincipalAvecPontsTexte = ""
    TempsAnodisationAvecPontsTexte = ""
    TempsMouvementsApresPostePrincipalTexte = ""
    TempsApresPostePrincipalAvecPontsTexte = ""
    TempsTotalPostesAvecPontsTexte = ""
    TempsTotalEgouttagesAvecPontsTexte = ""
    TempsTotalMouvementsTexte = ""
    TempsTotalGammeAvecPontsTexte = ""

    '--- transfert des d�tails de la gamme dans une gamme virtuel ---
    For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
        With TGammesAnodisation.TDetailsGammesAnodisation(a)
            .NumZone = TDetailsGammesAnodisation(a).NumZone
            
            .TempsAuPosteSecondes = TDetailsGammesAnodisation(a).TempsAuPosteSecondes
            .TempsAuPosteTexte = TDetailsGammesAnodisation(a).TempsAuPosteTexte
            
            .TempsAlerteSecondes = TDetailsGammesAnodisation(a).TempsAlerteSecondes
            .TempsAlerteTexte = TDetailsGammesAnodisation(a).TempsAlerteTexte
            
            .TempsEgouttageSecondes = TDetailsGammesAnodisation(a).TempsEgouttageSecondes
            .TempsEgouttageTexte = TDetailsGammesAnodisation(a).TempsEgouttageTexte
        
        End With
    Next a
    
    '--- calcul des temps avec les ponts ---
    CalculTempsGammeAnodisationAvecPonts TGammesAnodisation, _
                                                                          TempsMouvementsAvantPostePrincipalSecondes, _
                                                                          TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                          TempsPostePrincipalAvecPontsSecondes, _
                                                                          TempsMouvementsApresPostePrincipalSecondes, _
                                                                          TempsApresPostePrincipalAvecPontsSecondes, _
                                                                          TempsTotalPostesAvecPontsSecondes, _
                                                                          TempsTotalEgouttagesAvecPontsSecondes, _
                                                                          TempsTotalMouvementsSecondes, _
                                                                          TempsTotalGammeAvecPontsSecondes


    '--- affichage du temps des mouvements avant anodisation ---
    If TempsMouvementsAvantPostePrincipalSecondes = 0 Then
        TempsMouvementsAvantPostePrincipalTexte = PAS_DE_TEMPS
    Else
        TempsMouvementsAvantPostePrincipalTexte = CTemps2(TempsMouvementsAvantPostePrincipalSecondes)
    End If
    If LTempsMouvementsAvantPostePrincipal.Caption <> TempsMouvementsAvantPostePrincipalTexte Then
        LTempsMouvementsAvantPostePrincipal.Caption = TempsMouvementsAvantPostePrincipalTexte
    End If
    
    '--- affichage du temps avant anodisation avec les ponts ---
    If TempsAvantPostePrincipalAvecPontsSecondes = 0 Then
        TempsAvantPostePrincipalAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsAvantPostePrincipalAvecPontsTexte = CTemps2(TempsAvantPostePrincipalAvecPontsSecondes)
    End If
    If LTempsAvantPostePrincipalAvecPonts.Caption <> TempsAvantPostePrincipalAvecPontsTexte Then
        LTempsAvantPostePrincipalAvecPonts.Caption = TempsAvantPostePrincipalAvecPontsTexte
    End If
    
    '--- affichage du temps au poste d'anodisation (identique aux valeurs d�finies dans la gamme) ---
    If TempsPostePrincipalAvecPontsSecondes = 0 Then
        TempsAnodisationAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsAnodisationAvecPontsTexte = CTemps2(TempsPostePrincipalAvecPontsSecondes)
    End If
    If LTempsPostePrincipalAvecPonts.Caption <> TempsAnodisationAvecPontsTexte Then
        LTempsPostePrincipalAvecPonts.Caption = TempsAnodisationAvecPontsTexte
    End If
    
    '--- affichage du temps des mouvements apr�s anodisation ---
    If TempsMouvementsApresPostePrincipalSecondes = 0 Then
        TempsMouvementsApresPostePrincipalTexte = PAS_DE_TEMPS
    Else
        TempsMouvementsApresPostePrincipalTexte = CTemps2(TempsMouvementsApresPostePrincipalSecondes)
    End If
    If LTempsMouvementsApresPostePrincipal.Caption <> TempsMouvementsApresPostePrincipalTexte Then
        LTempsMouvementsApresPostePrincipal.Caption = TempsMouvementsApresPostePrincipalTexte
    End If
    
    '--- affichage du temps apr�s anodisation ---
    If TempsApresPostePrincipalAvecPontsSecondes = 0 Then
        TempsApresPostePrincipalAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsApresPostePrincipalAvecPontsTexte = CTemps2(TempsApresPostePrincipalAvecPontsSecondes)
    End If
    If LTempsApresPostePrincipalAvecPonts.Caption <> TempsApresPostePrincipalAvecPontsTexte Then
        LTempsApresPostePrincipalAvecPonts.Caption = TempsApresPostePrincipalAvecPontsTexte
    End If
    
    '--- affichage du temps total de la gamme ---
    If TempsTotalGammeAvecPontsSecondes = 0 Then
        TempsTotalGammeAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalGammeAvecPontsTexte = CTemps2(TempsTotalGammeAvecPontsSecondes)
    End If
    If LTempsTotalGammeAvecPonts.Caption <> TempsTotalGammeAvecPontsTexte Then
        LTempsTotalGammeAvecPonts.Caption = TempsTotalGammeAvecPontsTexte
    End If
    
    '--- affichage du temps total des mouvements ---
    If TempsTotalMouvementsSecondes = 0 Then
        TempsTotalMouvementsTexte = PAS_DE_TEMPS
    Else
        TempsTotalMouvementsTexte = CTemps2(TempsTotalMouvementsSecondes)
    End If
    If LTempsTotalMouvements.Caption <> TempsTotalMouvementsTexte Then
        LTempsTotalMouvements.Caption = TempsTotalMouvementsTexte
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la simulation de l'entr�e d'une charge dans la ligne
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SimulationEntreeCharge()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affichage de la simulation de l'entr�e d'une charge en ligne ---
    If TempsTotalGammeAvecPontsSecondes = 0 Then
        LSimulationEntreeCharge(0).Caption = "La simulation de l'entr�e d'une charge n'est pas possible pour le moment"
        LSimulationEntreeCharge(1).Caption = ""
    Else
        LSimulationEntreeCharge(0).Caption = "Entr�e d'une charge MAINTENANT (" & Format(Now, "hh:mm") & ")"
        LSimulationEntreeCharge(1).Caption = "Sortie de la ligne pr�vue vers " & _
                                                                       Format(DateAdd("s", TempsTotalGammeAvecPontsSecondes, Now), "hh:mm")
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affichage des temps principaux de la gamme d'anodisation
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageTempsGamme()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim a As Integer, _
           NumZone As Integer
    Dim TempsTexte As String
    Dim TGammesAnodisation As EnrGammesAnodisation
    
    '--- initialisation des variables ---
    TempsAvantPostePrincipalSansPontsSecondes = 0
    TempsPostePrincipalSansPontsSecondes = 0
    TempsApresPostePrincipalSansPontsSecondes = 0
    TempsTotalPostesSansPontsSecondes = 0
    TempsTotalEgouttagesSansPontsSecondes = 0
    TempsTotalGammeSansPontsSecondes = 0
    
    TempsAvantPostePrincipalSansPontsTexte = ""
    TempsPostePrincipalSansPontsTexte = ""
    TempsApresPostePrincipalSansPontsTexte = ""
    TempsTotalPostesSansPontsTexte = ""
    TempsTotalEgouttagesSansPontsTexte = ""
    TempsTotalGammeSansPontsTexte = ""
    
    '--- transfert des d�tails de la gamme dans une gamme virtuel ---
    For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
        With TGammesAnodisation.TDetailsGammesAnodisation(a)
            
            .NumZone = TDetailsGammesAnodisation(a).NumZone
            
            .TempsAuPosteSecondes = TDetailsGammesAnodisation(a).TempsAuPosteSecondes
            .TempsAuPosteTexte = TDetailsGammesAnodisation(a).TempsAuPosteTexte
            
            .TempsAlerteSecondes = TDetailsGammesAnodisation(a).TempsAlerteSecondes
            .TempsAlerteTexte = TDetailsGammesAnodisation(a).TempsAlerteTexte
            
            .TempsEgouttageSecondes = TDetailsGammesAnodisation(a).TempsEgouttageSecondes
            .TempsEgouttageTexte = TDetailsGammesAnodisation(a).TempsEgouttageTexte
        
        End With
    Next a
    
    '--- calcul des temps avec les ponts ---
    CalculTempsGammeAnodisationSansPonts TGammesAnodisation, _
                                                                           TempsAvantPostePrincipalSansPontsSecondes, _
                                                                           TempsPostePrincipalSansPontsSecondes, _
                                                                           TempsApresPostePrincipalSansPontsSecondes, _
                                                                           TempsTotalPostesSansPontsSecondes, _
                                                                           TempsTotalEgouttagesSansPontsSecondes, _
                                                                           TempsTotalGammeSansPontsSecondes
    
    '--- affichage du temps avant anodisation ---
    If TempsAvantPostePrincipalSansPontsSecondes = 0 Then
        TempsAvantPostePrincipalSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsAvantPostePrincipalSansPontsTexte = CTemps2(TempsAvantPostePrincipalSansPontsSecondes)
    End If
    For a = LTempsAvantPostePrincipalSansPonts.LBound To LTempsAvantPostePrincipalSansPonts.UBound
        If LTempsAvantPostePrincipalSansPonts(a).Caption <> TempsAvantPostePrincipalSansPontsTexte Then
            LTempsAvantPostePrincipalSansPonts(a).Caption = TempsAvantPostePrincipalSansPontsTexte
        End If
    Next a
    
    '--- affichage du temps au poste d'anodisation ---
    If TempsPostePrincipalSansPontsSecondes = 0 Then
        TempsPostePrincipalSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsPostePrincipalSansPontsTexte = CTemps2(TempsPostePrincipalSansPontsSecondes)
    End If
    For a = LTempsPostePrincipalSansPonts.LBound To LTempsPostePrincipalSansPonts.UBound
        If LTempsPostePrincipalSansPonts(a).Caption <> TempsPostePrincipalSansPontsTexte Then
            LTempsPostePrincipalSansPonts(a).Caption = TempsPostePrincipalSansPontsTexte
        End If
    Next a

    '--- affichage du temps apr�s anodisation ---
    If TempsApresPostePrincipalSansPontsSecondes = 0 Then
        TempsApresPostePrincipalSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsApresPostePrincipalSansPontsTexte = CTemps2(TempsApresPostePrincipalSansPontsSecondes)
    End If
    For a = LTempsApresPostePrincipalSansPonts.LBound To LTempsApresPostePrincipalSansPonts.UBound
        If LTempsApresPostePrincipalSansPonts(a).Caption <> TempsApresPostePrincipalSansPontsTexte Then
            LTempsApresPostePrincipalSansPonts(a).Caption = TempsApresPostePrincipalSansPontsTexte
        End If
    Next a

    '--- affectation du temps total des postes en texte ---
    If TempsTotalPostesSansPontsSecondes = 0 Then
        TempsTotalPostesSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalPostesSansPontsTexte = CTemps2(TempsTotalPostesSansPontsSecondes)
    End If
                
    '--- affectation du temps total des �gouttages en texte ---
    If TempsTotalEgouttagesSansPontsSecondes = 0 Then
        TempsTotalEgouttagesSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalEgouttagesSansPontsTexte = CTemps2(TempsTotalEgouttagesSansPontsSecondes)
    End If
    
    '--- affichage du temps total de la gamme ---
    If TempsTotalGammeSansPontsSecondes = 0 Then
        TempsTotalGammeSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalGammeSansPontsTexte = CTemps2(TempsTotalGammeSansPontsSecondes)
    End If
    For a = LTempsTotalGammeSansPonts.LBound To LTempsTotalGammeSansPonts.UBound
        If LTempsTotalGammeSansPonts(a).Caption <> TempsTotalGammeSansPontsTexte Then
            LTempsTotalGammeSansPonts(a).Caption = TempsTotalGammeSansPontsTexte
        End If
    Next a

End Sub

Private Sub TBNumGamme_Change()
    On Error Resume Next
    With TBNumGamme
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub TBNumGamme_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBNumGamme_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6
End Sub

Private Sub TBNouveauNumGamme_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6
End Sub

Private Sub TBNouveauNumGamme_LostFocus()
    On Error Resume Next
    If TBNouveauNumGamme.Text <> "" Then
        TBNouveauNumGamme.Text = Right(String(6, "0") & TBNouveauNumGamme.Text, 6)
    End If
End Sub

Private Sub TBNumGamme_LostFocus()
    On Error Resume Next
    If TBNumGamme.Text <> "" Then
        TBNumGamme.Text = Right(String(6, "0") & TBNumGamme.Text, 6)
    End If
End Sub

Private Sub TBRefGamme_Change()
    On Error Resume Next
    With TBRefGamme
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub TBRefGamme_GotFocus()
    On Error Resume Next
    Me.ActiveControl.MaxLength = ADODCGammesAnodisation.Recordset(Me.ActiveControl.DataField).DefinedSize
End Sub

Private Sub TBRefGamme_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBRefGamme_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE
End Sub

Private Sub TBTensionsPhases_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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
        .SelStart = 0          'met en surbrillance la s�lection saisie
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
    GestionBoutons E_MODIFICATION_EN_COURS
End Sub

Private Sub TBTensionsPhases_LostFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
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

Private Sub TDBGGrilleRecherche_Click()
    On Error Resume Next
    If Me.ActiveControl.Name <> TDBGGrilleRecherche.Name Then           'placer le focus sur la grille si n�cessaire
        TDBGGrilleRecherche.SetFocus
    End If
End Sub

Private Sub TDBGGrilleRecherche_DblClick()
    On Error Resume Next
    CBRechercherSurGrille_Click
End Sub

Private Sub TDBGGrilleRecherche_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = vbDataErrContinue
End Sub

Private Sub TDBGGrilleRecherche_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�placement du focus sur le bouton ---
    With SFocusGrilleRecherche
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height + Screen.TwipsPerPixelY
        .Width = ActiveControl.Width + Screen.TwipsPerPixelX
        .Visible = True
    End With

    '--- affichage de la barre de s�lection ---
    With TDBGGrilleRecherche
        .CurrentCellVisible = True
        .Refresh
    End With

End Sub

Private Sub TDBGGrilleRecherche_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine ---
    Select Case KeyCode
        Case vbKeyReturn
            CBRechercherSurGrille_Click
            KeyCode = 0: Shift = 0
        Case vbKeyHome
            If Shift = vbCtrlMask Then
                ADODCGammesAnodisation.Recordset.MoveFirst
                KeyCode = 0: Shift = 0
            End If
        Case vbKeyEnd
            If Shift = vbCtrlMask Then
                ADODCGammesAnodisation.Recordset.MoveLast
                KeyCode = 0: Shift = 0
            End If
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown
        Case vbKeyTab
            If Shift = vbShiftMask Then
                TBContenant.SetFocus
            End If
            KeyCode = 0: Shift = 0
        Case Else: KeyCode = 0: Shift = 0
    End Select

End Sub

Private Sub TDBGGrilleRecherche_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- cadre de focus ---
    SFocusGrilleRecherche.Visible = False

End Sub

Private Sub TimerSimulationEntreeCharge_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerSimulationEntreeCharge.Enabled = False
    SimulationEntreeCharge
    TimerSimulationEntreeCharge.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Insertion d'une mati�re
' Entr�es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InsertionMatiere()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim a As Integer                            'pour les boucles FOR...NEXT

    '--- lancer la modification ---
    GestionBoutons E_MODIFICATION_EN_COURS

    '--- recherche du premier champ vide et affectation ---
    For a = 1 To NBR_MATIERES_MAXI_PAR_GAMME
    
        If TBMatieres(a).Text = "" Then
            ADODCGammesAnodisation.Recordset("Matiere" & a).value = ADODCMatieres.Recordset("Matiere").value
            Exit For
        End If
    
    Next a

End Sub

Private Sub TOBGestionGrilles_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer                                                      'pour les boucles FOR...NEXT
    Dim NumLigne As Integer                                        'num�ro de ligne
    Dim FicheVide As ImgDetailsGammesProduction   'fiche vide � l'image des gammes de production
    
    '--- affectation ---

    '--- s�lection en fonction de l'outil cliqu� ---
    Select Case Button.Key

        Case "SupprimerLigne"
            '--- supprimer une ligne ---
            NumLigne = MSHFGDetailsGammesAnodisation.Row
    
            '--- suppression de la ligne ---
            If NumLigne > 0 And NumLigne <= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                If AppelFenetre(F_MESSAGE, _
                                        TITRE_MESSAGES, _
                                        MESSAGE_3 & CStr(NumLigne) & " ?", _
                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                        TYPES_BOUTONS.T_OUI_NON, _
                                        EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                    TDetailsGammesAnodisation(NumLigne) = FicheVide
                    GestionDetailsGammesAnodisation GG_COMPRESSION
                    GestionDetailsGammesAnodisation GG_AFFICHAGE
                    GestionBoutons E_MODIFICATION_EN_COURS
                End If
                MSHFGDetailsGammesAnodisation.SetFocus
            End If
        
        Case "CompacterGrille"
            '--- compacter la grille ---
            GestionDetailsGammesAnodisation GG_COMPRESSION
            GestionDetailsGammesAnodisation GG_AFFICHAGE
        
        Case "InsererLigne"
            '--- ins�rer ligne ---
            NumLigne = MSHFGDetailsGammesAnodisation.Row

            '--- insertion de la ligne ---
            If NumLigne > 0 And NumLigne <= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                For a = Pred(NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) To NumLigne Step -1
                    TDetailsGammesAnodisation(Succ(a)) = TDetailsGammesAnodisation(a)
                Next a
                TDetailsGammesAnodisation(NumLigne) = FicheVide
                GestionDetailsGammesAnodisation GG_AFFICHAGE
                With MSHFGDetailsGammesAnodisation
                    .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                    .SetFocus
                End With
            End If
            
        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Gestion de la grille de recherche
' Entr�es : EtatSouhaite -> Fonction de l'�num�ration GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionGrilleRecherche(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    
    '--- affectation ---

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGGrilleRecherche
                
                .Visible = False                                                            'rendre la grille invisible
                '.ClearFields                                                                  'effacer la structure
             
                .Splits(0).AllowSizing = True                                        'autorise le fractionnement de la grille (petite rectangle noir en bas � gauche)
            
                .HeadLines = 3                                                             'nombre de ligne des ent�tes
                .HeadBackColor = COULEURS.BLEU_5                      'couleur de fond des ent�tes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des ent�tes
                
                .DeadAreaBackColor = COULEURS.JAUNE_0              'couleur de la surface non utilis�e
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.ORANGE_1  'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                .SelectedBackColor = COULEURS.ROUGE_3                'couleur de fond pour la s�lection
                .SelectedForeColor = COULEURS.JAUNE_3                  'couleur de premier plan pour la s�lection
                
                .HeadFont.Name = "Arial"
                With .Font
                    .Name = "MS Sans serif"
                    .Bold = True                                                              'caract�res gras
                End With
                
                .RowHeight = 0                                                              '�paisseur des lignes
                .RowHeight = .RowHeight * 1.05
                
                .RecordSelectors = True                                                 'affichage du s�lecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 '�paisseur du s�lecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du s�lecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du s�lecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = Me.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = Me.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = Me.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = False                                                  'interdire un nouvel enregistrement
                .AllowDelete = False                                                     'interdire la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire la s�lection des colonnes
                .AllowColMove = False                                                 'interdire le d�placement des colonnes s�lectionn�es
                
                .AllowRowSelect = True                                                'autoriser la s�lection des lignes
                .AllowRowSizing = True                                                'autoriser la modification de l'�paisseur des lignes
                
                .DataView = dbgNormalView                                         'pr�sentation normale de la grille
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGGrilleRecherche
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_GAMME)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N� de gamme"
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
                    .Caption = "R�f�rence de la gamme"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_DATE_CREATION_GAMME)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Date de cr�ation"
                    .Width = EPAISSEUR_CARACTERE * 17
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
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

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_MATIERE_1)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Mati�re concern�e 1"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_MATIERE_2)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Mati�re concern�e 2"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_MATIERE_3)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Mati�re concern�e 3"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_MATIERE_4)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Mati�re concern�e 4"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_MATIERE_5)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Mati�re concern�e 5"
                    .Width = EPAISSEUR_CARACTERE * 20
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
' R�le      : Initialise les champs de la partie redresseur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationChampsRedresseur()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const INIT_TEMPS As String = "0:00:00"
    Const INIT_TENSION As String = "0,0"
    Const INIT_INTENSITE As String = "0"

    '--- d�claration ---
    Dim a As Integer                                            'pour les boucles FOR...NEXT

    '--- affectation ---
    
    '--- interdire les �v�nements ---
    InterdireEvenements = True

    '--- for�age du mode U ou I en mode tension ---
    Call LModeUouI_Click(MODES_U_OU_I.M_TENSION)

    '--- initialisation des champs temps, tension, intensit� ---
    For a = MEBTempsPhases.LBound To MEBTempsPhases.UBound
        MEBTempsPhases(a).Text = INIT_TEMPS
        TBTensionsPhases(a).Text = INIT_TENSION
        TBIntensitesPhases(a).Text = INIT_INTENSITE
    Next a

    '--- temps total de la gamme redresseur ---
    LTempsTotalGammeRedresseur.Caption = INIT_TEMPS

    '--- autorisation des �v�nements ---
    InterdireEvenements = False
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Calcul du temps total de la gamme redresseur
' Entr�es : CalculTempsTotalGammeRedresseur -> Le temps total de la gamme en secondes
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CalculTempsTotalGammeRedresseur() As Long
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer                            'pour les boucles FOR...NEXT

    '--- affectation ---
    CalculTempsTotalGammeRedresseur = 0
    
    '--- calcul du temps ---
    For a = MEBTempsPhases.LBound To MEBTempsPhases.UBound
            CalculTempsTotalGammeRedresseur = CalculTempsTotalGammeRedresseur + CTempsTexteEnSecondes(MEBTempsPhases(a).Text)
    Next a

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Recherche le passage dans les postes principaux
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub RecherchePassageBainsPrincipaux(ByRef PassageAnodisation As Boolean, _
                                                                             ByRef PassageSpectro As Boolean, _
                                                                             ByRef PassageOr As Boolean, _
                                                                             ByRef PassageNoir As Boolean)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---

    '--- d�claration ---
    Dim a As Integer                                                                  'r�serv� pour les boucles FOR ... NEXT
    Dim Codezone As String                                                      'code de la zone

    '--- affectation par d�faut ---
    PassageAnodisation = False
    PassageSpectro = False
    PassageOr = False
    PassageNoir = False
    
    '--- affectation par d�faut ---
    For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
    
        With TDetailsGammesAnodisation(a)
                        
            Select Case Trim(.Codezone)
            
                Case "": Exit For
                Case "C13 � C16": PassageAnodisation = True
                Case "C19": PassageSpectro = True
                Case "C22": PassageOr = True
                Case "C28": PassageNoir = True
                
                Case Else
            End Select
            
        End With
    
    Next a

End Sub


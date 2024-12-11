VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FPremisses 
   ClientHeight    =   9825
   ClientLeft      =   1245
   ClientTop       =   630
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   10215
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
      Height          =   1095
      Left            =   240
      ScaleHeight     =   1065
      ScaleWidth      =   28185
      TabIndex        =   11
      Top             =   660
      Width           =   28215
      Begin VB.ComboBox CBNumPosteArrivee 
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
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   5895
      End
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
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   480
         Width           =   5895
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
         TabIndex        =   15
         ToolTipText     =   " Change la présentation de la grille "
         Top             =   540
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
         TabIndex        =   14
         Top             =   180
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
         Left            =   1740
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   " Annule tris et recherches "
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CBRechercherSurGrille 
         BackColor       =   &H00E0E0E0&
         Caption         =   "GRILLE"
         DownPicture     =   "FPremisses.frx":01F2
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
         Picture         =   "FPremisses.frx":08F4
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   " Rechercher sur la grille "
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin TrueOleDBGrid80.TDBGrid TDBGGrilleRecherche 
         Bindings        =   "FPremisses.frx":0FF6
         Height          =   10875
         Left            =   240
         TabIndex        =   16
         Top             =   1140
         Width           =   27675
         _ExtentX        =   48816
         _ExtentY        =   19182
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NumPosteDepart"
         Columns(0).DataField=   "NumPosteDepart"
         Columns(0).DataWidth=   6
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NomPosteDepart"
         Columns(1).DataField=   "NomPosteDepart"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NumPosteArrivee"
         Columns(2).DataField=   "NumPosteArrivee"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "NomPosteArrivee"
         Columns(3).DataField=   "NomPosteArrivee"
         Columns(3).DataWidth=   6
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "NumPont"
         Columns(4).DataField=   "NumPont"
         Columns(4).DataWidth=   6
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "PremisseCodee"
         Columns(5).DataField=   "PremisseCodee"
         Columns(5).DataWidth=   200
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "PremisseDecodee"
         Columns(6).DataField=   "PremisseDecodee"
         Columns(6).DataWidth=   200
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   -1  'True
         Splits(0)._GSX_SAVERECORDSELECTORS=   0
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4366"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4233"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4366"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=3572"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3440"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1984"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1852"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(4)._AlignLeft=0"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=4366"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=4233"
         Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(26)=   "Column(6).Width=4366"
         Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=4233"
         Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
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
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Poste d'ARRIVEE"
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
         Index           =   0
         Left            =   8985
         TabIndex        =   20
         Top             =   180
         Width           =   5940
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
         BackStyle       =   0  'Transparent
         Caption         =   "Poste de DEPART"
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
         Index           =   11
         Left            =   2880
         TabIndex        =   17
         Top             =   180
         Width           =   5910
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FPremisses.frx":1013
      ScaleHeight     =   315
      ScaleWidth      =   10155
      TabIndex        =   9
      Top             =   0
      Width           =   10215
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "PREMISSE GEREE"
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
         TabIndex        =   10
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
      ScaleWidth      =   10155
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   8730
      Width           =   10215
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FPremisses.frx":25955
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
         Left            =   21480
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":26057
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBRegenerationComplete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Régénération COMPLETE"
         DownPicture     =   "FPremisses.frx":26759
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
         Left            =   5220
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":27BDB
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Régénération complète de toutes les prémisses "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3255
      End
      Begin VB.CommandButton CBSupprimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Supprimer"
         DownPicture     =   "FPremisses.frx":2905D
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
         Left            =   12180
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":2975F
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Supprimer l'enregistrement en cours "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBPremisseAutomatique 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Prémisse AUTOMATIQUE"
         DownPicture     =   "FPremisses.frx":29E61
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
         Left            =   8700
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":2B2E3
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Génération de la prémisse en automatique "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3255
      End
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FPremisses.frx":2C765
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
         Height          =   855
         Left            =   18240
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":2CE67
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Annuler les dernières modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FPremisses.frx":2D569
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
         Height          =   855
         Left            =   19860
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":2DC6B
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualise&r"
         DownPicture     =   "FPremisses.frx":2E36D
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
         Left            =   13860
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPremisses.frx":2EA6F
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   " Actualiser les données "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin MSAdodcLib.Adodc ADODCPremisses 
         Height          =   435
         Left            =   16680
         Top             =   540
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   767
         ConnectMode     =   1
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   1
         CommandType     =   8
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
         RecordSource    =   $"FPremisses.frx":2F171
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
      Begin MSComctlLib.ImageList ILGrillesDonnees 
         Left            =   2760
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
               Picture         =   "FPremisses.frx":2F332
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":2F53E
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":2F74A
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":2F956
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":2FB62
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":2FD6E
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":2FF7A
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":30186
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":30392
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":3059E
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":307AA
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":309B6
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":30BC2
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":30DCE
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":30FDA
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":311E6
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":313F2
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":315F6
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":317FA
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":319FE
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":31C02
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":31E06
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":3200A
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":3220E
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":32412
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":32616
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ILOutilsGestionGrilles 
         Left            =   3420
         Top             =   120
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
               Picture         =   "FPremisses.frx":3281A
               Key             =   "supprimer"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":33C04
               Key             =   "compacter"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FPremisses.frx":34FEE
               Key             =   "inserer"
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   315
         Left            =   4140
         Top             =   180
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
         Left            =   16680
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox PBPremisseGeree 
      BackColor       =   &H00FFFFFF&
      Height          =   10875
      Left            =   240
      ScaleHeight     =   10815
      ScaleWidth      =   28095
      TabIndex        =   21
      Top             =   2040
      Width           =   28155
      Begin VB.Frame FChoixIA 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Choix du moteur d'inférence "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   21540
         TabIndex        =   32
         Top             =   1860
         Width           =   6315
         Begin VB.Label LTempsCycleSecondes 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   300
            TabIndex        =   36
            Top             =   1140
            Width           =   5715
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Temps du cycle en secondes (total du temps de chaque action)"
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
            Index           =   6
            Left            =   105
            TabIndex        =   35
            Top             =   780
            Width           =   6045
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pont "
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
            Index           =   7
            Left            =   480
            TabIndex        =   34
            Top             =   420
            Width           =   1215
            WordWrap        =   -1  'True
         End
         Begin VB.Label LNomPontIA 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
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
            Left            =   1800
            TabIndex        =   33
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.Frame FChoixBase 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Choix de base "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   21540
         TabIndex        =   26
         Top             =   240
         Width           =   6315
         Begin VB.ComboBox ComboPont 
            Height          =   315
            Left            =   1800
            TabIndex        =   38
            Text            =   "Combo1"
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pont concerné"
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
            Left            =   -60
            TabIndex        =   31
            Top             =   900
            Width           =   1725
            WordWrap        =   -1  'True
         End
         Begin VB.Label LNomPosteArrivee 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "NomPosteArrivee"
            DataSource      =   "ADODCPremisses"
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
            Left            =   4800
            TabIndex        =   30
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Poste d'arrivée"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   3120
            TabIndex        =   29
            Top             =   360
            Width           =   1575
            WordWrap        =   -1  'True
         End
         Begin VB.Label LNomPosteDepart 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            BorderStyle     =   1  'Fixed Single
            DataField       =   "NomPosteDepart"
            DataSource      =   "ADODCPremisses"
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
            Left            =   1800
            TabIndex        =   28
            Top             =   300
            Width           =   1275
         End
         Begin VB.Label LLibelles 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Poste de départ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   1635
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame FDetailsPremisses 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Détails de la prémisse "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   10335
         Left            =   300
         TabIndex        =   22
         Top             =   240
         Width           =   21015
         Begin MSMask.MaskEdBox MEBEditionDetailsPremisses 
            Height          =   255
            Left            =   6960
            TabIndex        =   23
            Top             =   180
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            BackColor       =   12632319
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
         Begin MSAdodcLib.Adodc ADODCActions 
            Height          =   375
            Left            =   300
            Top             =   9660
            Width           =   9195
            _ExtentX        =   16219
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
            Connect         =   "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
            OLEDBString     =   "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   $"FPremisses.frx":363D8
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
         Begin MSDataGridLib.DataGrid DGActions 
            Bindings        =   "FPremisses.frx":3640D
            Height          =   8955
            Left            =   300
            TabIndex        =   24
            Top             =   420
            Width           =   9195
            _ExtentX        =   16219
            _ExtentY        =   15796
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
            ColumnCount     =   3
            BeginProperty Column00 
               DataField       =   "CodeAction"
               Caption         =   "Code de l'action"
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
               DataField       =   "LibelleAction"
               Caption         =   "Libellé de l'action"
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
               DataField       =   "NumAction"
               Caption         =   "N° de l'action"
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
                  Alignment       =   2
                  DividerStyle    =   3
                  ColumnAllowSizing=   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   2174,74
               EndProperty
               BeginProperty Column01 
                  DividerStyle    =   3
                  ColumnAllowSizing=   -1  'True
                  Locked          =   -1  'True
                  ColumnWidth     =   6765,166
               EndProperty
               BeginProperty Column02 
               EndProperty
            EndProperty
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsPremisses 
            Height          =   8925
            Left            =   9600
            TabIndex        =   25
            Top             =   1080
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   15743
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   12582912
            Rows            =   31
            Cols            =   6
            BackColorFixed  =   12582912
            ForeColorFixed  =   16777215
            BackColorBkg    =   12648447
            GridColor       =   12632256
            GridColorUnpopulated=   -2147483644
            AllowBigSelection=   0   'False
            FocusRect       =   2
            HighLight       =   0
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
         Begin MSComctlLib.Toolbar TOBGestionGrilles 
            Height          =   405
            Left            =   9780
            TabIndex        =   37
            Top             =   420
            Width           =   10950
            _ExtentX        =   19315
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
                  Object.ToolTipText     =   " Insère une ligne dans une grille "
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Shape SFocusTableDetailsPremisses 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Height          =   8940
            Left            =   9780
            Top             =   1080
            Visible         =   0   'False
            Width           =   10950
         End
         Begin VB.Shape SFocusTableCodesActions 
            BorderColor     =   &H000000FF&
            BorderWidth     =   4
            Height          =   8970
            Left            =   300
            Top             =   420
            Visible         =   0   'False
            Width           =   9210
         End
      End
   End
End
Attribute VB_Name = "FPremisses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les prémisses
' Nom                    : FPremisses.frm
' Date de création : 20/12/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const NBR_COLONNES_DETAILS_PREMISSES  As Integer = 3

Private Const TITRE_FENETRE As String = "PREMISSES"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum COLONNES_DETAILS_PREMISSES
    C_NUM_LIGNES = 0
    C_CODE_ACTION = 1
    C_PARAMETRE = 2
    C_LIBELLE_ACTION = 3
End Enum

Private Enum COLONNES_GRILLE_RECHERCHE
    
    C_NUM_POSTE_DEPART = 0
    C_NOM_POSTE_DEPART = 1
    
    C_NUM_POSTE_ARRIVEE = 2
    C_NOM_POSTE_ARRIVEE = 3
    
    C_NUM_PONT = 4
    
    C_PREMISSE_CODEE = 5
    C_PREMISSE_DECODEE = 6

End Enum

'--- types privés ---
Private Type ImgDetailsPremisses
    NumAction As Integer                        'n° de l'action
    CodeAction As String                         'code de l'action
    ParametreOuiNon As Boolean          'paramètre oui ou non
    Parametre As String                          'paramètre en fonction de l'action
    LibelleAction As String                      'libellé de l'action
End Type

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean        'pour interdire certains évènements
Private LigneDepartDeplacement As Integer   'ligne de départ en cas de déplacement d'un détail
Private LigneArriveeDeplacement As Integer  'ligne de d'arrivée en cas de déplacement d'un détail
Private MemDernierBouton As Long                'mémoire du dernier bouton

'--- tableaux privés ---
Private TDetailsPremisses(1 To NBR_LIGNES_DETAILS_PREMISSES) As ImgDetailsPremisses

'--- variables publiques ---
Public RechercherSurGrille As Boolean          'publique pour le copier / coller
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active

Private Sub ADODCActions_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
    With pRecordset
        If .BOF = False And .EOF = False Then
            ADODCActions.Caption = .Fields("CodeAction") & " - " & .Fields("LibelleAction")
        End If
    End With

End Sub

Private Sub ADODCPremisses_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---

    With pRecordset
        
        If .BOF = False And .EOF = False Then
        
            '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
            If pRecordset("NumPosteDepart") >= POSTES.P_CHGT_1 And pRecordset("NumPosteDepart") <= DERNIER_POSTE And _
               pRecordset("NumPosteArrivee") >= POSTES.P_CHGT_1 And pRecordset("NumPosteArrivee") <= DERNIER_POSTE Then
                
                Me.Caption = TITRE_FENETRE & " - " & _
                                       "Poste de départ " & _
                                       TEtatsPostes(pRecordset("NumPosteDepart")).DefinitionPoste.NomPoste & _
                                        ", poste d'arrivée " & _
                                       TEtatsPostes(pRecordset("NumPosteArrivee")).DefinitionPoste.NomPoste
            
            End If
            LRenseignements.Caption = .AbsolutePosition & "/" & .RecordCount
       
        Else
               
            '--- si fiche inexistante affichage d'un tiret ---
            Me.Caption = TITRE_FENETRE
            LRenseignements.Caption = "-"
        
        End If
    
        '--- affichage des renseignements de la fenetre ---
        LRenseignementsFenetre.Caption = Me.Caption
    
    End With
    
    '--- chargement des détails ---
    If PremiereActivation = True And RechercherSurGrille = False Then
        LectureDetailsPremisses
        LecturePartieIA
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le lien entre les données et les critères de recherche
' Entrées : NumPosteDepart -> Poste de départ
'                NumPosteArrivee -> Poste d'arrivée
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LieDonneesRecherche(ByVal NumPosteDepart As Integer, _
                                                         ByVal NumPosteArrivee As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
                
    '--- interdire les évènements ---
    InterdireEvenements = True
            
    '--- valeur des listes ---
    CBNumPosteDepart.ListIndex = NumPosteDepart
    CBNumPosteArrivee.ListIndex = NumPosteArrivee
    
    '--- autoriser certains évènements ---
    InterdireEvenements = False
                
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
    ADODCPremisses.Refresh
    
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

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_ANNULER
        
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- annuler ---
    ADODCPremisses.Recordset.CancelUpdate
    
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

Private Sub CBNumPosteArrivee_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- lancer le tri ---
    If InterdireEvenements = False Then
        LanceRechercheOuTri True
    End If

End Sub

Private Sub CBNumPosteDepart_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- lancer le tri ---
    If InterdireEvenements = False Then
        LanceRechercheOuTri True
    End If

End Sub

Private Sub CBPremisseAutomatique_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumPosteDepart As Integer, _
            NumPosteArrivee As Integer
    Dim PremisseDecodee As String
    
    '--- affectation ---
    With ADODCPremisses.Recordset
        NumPosteDepart = .Fields("NumPosteDepart")
        NumPosteArrivee = .Fields("NumPosteArrivee")
    End With
    
    '--- calcul automatique de la prémisse décodée ---
    PremisseDecodee = CalculAutomatiquePremisseDecodee(NumPosteDepart, NumPosteArrivee)
    'PremisseCodee = PremisseDecodeeVersCodee(PremisseDecodee)
    
    Dim ClePrimaire As Integer
    ClePrimaire = ADODCPremisses.Recordset.Fields("cleprimaire")
    
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- ouverture de la connexion à la base de données d'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseClient
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- affectation de la requête ---
    Requete = "UPDATE premisses SET PremisseDecodee='" & PremisseDecodee & _
                "' WHERE cleprimaire=" & ClePrimaire

    '--- lancement de la requête ---
    Enregistrement.Open Requete, ConnexionBDAnodisationSQL, adCmdText
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    ' *****************************************************************************
    
    Dim NumPont, NumPontIA    As Integer
    NumPontIA = CInt(LNomPontIA.Caption)
    If (ComboPont.ListIndex <> 0) Then
        NumPont = ComboPont.ListIndex
        LNomPontIA.Caption = NumPont
        NumPontIA = NumPont
    Else
        NumPont = NumPontIA
    End If
    
    '--- affectation ---
    With ADODCPremisses.Recordset
        NumPosteDepart = .Fields("NumPosteDepart")
        NumPosteArrivee = .Fields("NumPosteArrivee")
    End With
    

    
   With TPremisses(NumPosteDepart, NumPosteArrivee)
        .NumPont = NumPont
        .NumPontIA = NumPontIA
               
    End With
    
    
    Dim pos As Integer
    pos = ADODCPremisses.Recordset.AbsolutePosition
    
    
    ADODCPremisses.Refresh
    ADODCPremisses.Recordset.AbsolutePosition = pos
    'LNomPontIA.Caption = TPremisses(NumPosteDepart, NumPosteArrivee).NumPont
    '--- modification de l'enregistrement ---
    'ADODCPremisses.Recordset.Fields("PremisseDecodee") = PremisseDecodee

    '--- transfert de la prémisse dans le tableau ---
    GestionDetailsPremisses GG_VIDAGE
    GestionDetailsPremisses GG_AFFICHAGE
    GestionDetailsPremisses GG_TRANSFERT_DONNEES
    GestionDetailsPremisses GG_AFFICHAGE
    GestionBoutons E_MODIFICATION_EN_COURS

End Sub

Private Sub CBPremisseAutomatique_GotFocus()
    
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

Private Sub CBPremisseAutomatique_LostFocus()
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

Private Sub CBRaz_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    InterdireEvenements = True
    CBNumPosteDepart.ListIndex = 0
    CBNumPosteArrivee.ListIndex = 0
    InterdireEvenements = False
    
    '--- lancer le tri ---
    LanceRechercheOuTri True

End Sub

Private Sub CBRechercherSurGrille_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    If CBRechercherSurGrille.Enabled = True Then

        '--- affectation ---
        RechercherSurGrille = Not (RechercherSurGrille)
                
        '--- affichage ---
        AfficheGrilleRecherche
        
        '--- lancer la lecture des détails ---
        If PremiereActivation = True And RechercherSurGrille = False Then
            LectureDetailsPremisses
            LecturePartieIA
        End If
        
    End If

End Sub

Private Sub CBRegenerationComplete_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- demande de confirmation ---
    If AppelFenetre(F_MESSAGE, _
                             TITRE_MESSAGES, _
                             MESSAGE_601, _
                             TYPES_MESSAGES.T_AVERTISSEMENT, _
                             TYPES_BOUTONS.T_OUI_NON, _
                             EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
    
        '--- rafraichissement de l'écran ---
        Me.Refresh
    
        '--- lancement de la régénération ---
        Bidon = RegenerationCompletePremisses()
    
    End If

End Sub

Private Sub CBRegenerationComplete_GotFocus()
    
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

Private Sub CBRegenerationComplete_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBSupprimer_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_SUPPRIMER
    
    '--- demande de confirmation ---
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                            MESSAGE_600, _
                            TYPES_MESSAGES.T_AVERTISSEMENT, _
                            TYPES_BOUTONS.T_OUI_NON, _
                            EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
    
        '--- curseur de la souris ---
        SourisEnAttente True
        
        '--- ATTENTION effacement des données mais pas de l'enregistrement ---
        With ADODCPremisses.Recordset
            .Fields("PremisseCodee") = ""
            .Fields("PremisseDecodee") = ""
            .UpdateBatch adAffectAllChapters         'enregistrement
        End With
        
        '--- curseur de la souris ---
        SourisEnAttente False
    
        '--- actualiser l'enregistrement ---
        Me.Refresh
        CBActualiser_Click
    
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

    '--- déplacement du focus sur le bouton ---
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

Private Sub CBValider_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim DernierEtat As Integer
    Dim PremisseCodee As String, _
            PremisseDecodee As String

    '--- gestion des boutons ---
    DernierEtat = MemDernierBouton
    GestionBoutons E_AVANT_VALIDER
    
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- suppression et réenregistrement des détails ---
    
    '--- valeurs de champs ---
    ConstruitPremisses PremisseCodee, PremisseDecodee
    'With ADODCPremisses.Recordset
    '    .Fields("PremisseCodee") = PremisseCodee
    '    .Fields("PremisseDecodee") = PremisseDecodee
    'End With

    '--- valider l'enregistrement ---
    'ADODCPremisses.Recordset.UpdateBatch adAffectAllChapters
    
    '**********************************************************************
    '**********************************************************************
    
    Dim NumPont, NumPontIA, ClePrimaire As Integer
    NumPontIA = CInt(LNomPontIA.Caption)
    If (ComboPont.ListIndex <> 0) Then
        NumPont = ComboPont.ListIndex
        'LNomPontIA.Caption = NumPont
    Else
        NumPont = NumPontIA
    End If
    

    
    
    
    
    ClePrimaire = ADODCPremisses.Recordset.Fields("cleprimaire")
    
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- ouverture de la connexion à la base de données d'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseClient
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- affectation de la requête ---
    Requete = "UPDATE premisses SET PremisseCodee = '" & PremisseCodee & "' , PremisseDecodee='" & PremisseDecodee & _
                "',numpont=" & NumPont & ", numpontia=" & NumPontIA & " WHERE cleprimaire=" & ClePrimaire

    '--- lancement de la requête ---
    Enregistrement.Open Requete, ConnexionBDAnodisationSQL, adCmdText
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
    
    '**********************************************************************
    '**********************************************************************
    
    
    
    
    
    '--- actualisation ---
    Select Case DernierEtat
        Case ETATS_BOUTONS.E_APRES_NOUVEAU: ADODCPremisses.Recordset.Requery
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
        Case Else
    End Select
    
    '--- réactualisation des prémisses ---
    'ChargePremisses
    
    '--- curseur de la souris ---
    SourisEnAttente False
    
    Dim NumPosteDepart As Integer, _
            NumPosteArrivee As Integer

    
    '--- affectation ---
    With ADODCPremisses.Recordset
        NumPosteDepart = .Fields("NumPosteDepart")
        NumPosteArrivee = .Fields("NumPosteArrivee")
    End With
    

    
   With TPremisses(NumPosteDepart, NumPosteArrivee)
        .NumPont = NumPont
        .NumPontIA = NumPontIA
        .PremisseCodee = PremisseCodee
        .PremisseDecodee = PremisseDecodee
               
    End With
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

Private Sub ComboPont_Click()
    On Error Resume Next
    With ComboPont
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name And .DataChanged = True Then
                GestionBoutons E_MODIFICATION_EN_COURS
            End If
        End If
    End With
End Sub

Private Sub DCNumPont_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub DGActions_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn
            InsertionDetail 0
            KeyCode = 0: Shift = 0
        Case Else
    End Select
End Sub

Private Sub DGActions_DblClick()
    On Error Resume Next
    InsertionDetail
End Sub

Private Sub DGActions_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = vbDataErrContinue
End Sub

Private Sub DGActions_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- cadre de focus ---
    SFocusTableCodesActions.Visible = True

    '--- affichage de la barre de sélection ---
    With DGActions
        .CurrentCellVisible = True
        .Refresh
    End With

End Sub

Private Sub DGActions_HeadClick(ByVal ColIndex As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- lancement de la requête ---
    With ADODCActions

        '--- requête ---
        Select Case ColIndex
            Case 1
                .RecordSource = "SELECT Actions.* FROM Actions ORDER BY LibelleAction"
            Case Else
                .RecordSource = "SELECT Actions.* FROM Actions ORDER BY NumAction"
        End Select

        '--- rafraichissement ---
        .Refresh
        .Recordset.MoveFirst

    End With

    '--- effacement de la sélection de colonne ---
    DGActions.ClearSelCols

End Sub

Private Sub DGActions_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- cadre de focus ---
    SFocusTableCodesActions.Visible = False

End Sub

Private Sub DGActions_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    On Error Resume Next
    With DGActions 'pour fixer toujour la première colonne
        .Col = 0
        .CurrentCellVisible = True
    End With
End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        Me.Refresh
        LectureDetailsPremisses
        LecturePartieIA
        If CBRechercherSurGrille.Visible = True Then CBRechercherSurGrille.SetFocus
        PremiereActivation = True
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Marque ou restitue un enregistrement (fonction Bookmark)
' Entrées : MarquageRestitution -> TRUE  = Marquage de l'enregistrement
'                                                       FALSE = Restitution de l'enregistrement
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MarqueEnregistrement(ByVal MarquageRestitution As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Static SignetEnregistrement As Variant

    With ADODCPremisses.Recordset
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
' Rôle      : Lance une recherche en fonction des critères
' Entrées : MethodeChoisie -> FALSE = Recherche
'                                                TRUE  = Tri
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LanceRechercheOuTri(ByVal MethodeChoisie As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim ChaineDeRecherche As String, _
           RequeteSQL As String, _
           Filtre1 As String, _
           Filtre2 As String

    '--- curseur de la souris ---
    SourisEnAttente True

    If MethodeChoisie = True Then

        '--- début de la requête ---
        RequeteSQL = "SELECT premisses.cleprimaire,Postes1.NomPoste AS NomPosteDepart," & _
                                "Postes.NomPoste AS NomPosteArrivee, Premisses.NumPont, " & _
                                "Premisses.NumPosteDepart, " & _
                                "Premisses.NumPosteArrivee, Premisses.PremisseCodee, " & _
                                "Premisses.PremisseDecodee " & _
                                "FROM Premisses " & _
                                "LEFT OUTER JOIN Postes ON Premisses.NumPosteArrivee = Postes.NumPoste " & _
                                "LEFT OUTER JOIN Postes Postes1 ON Premisses.NumPosteDepart = Postes1.NumPoste "

        '--- construction du filtre ---
        With CBNumPosteDepart
            If .ListIndex > 0 Then
                Filtre1 = "Premisses.NumPosteDepart = " & .ItemData(.ListIndex)
            End If
        End With
        With CBNumPosteArrivee
            If .ListIndex > 0 Then
                Filtre2 = "Premisses.NumPosteArrivee = " & .ItemData(.ListIndex)
            End If
        End With
        If Filtre1 = "" And Filtre2 = "" Then
            Else
            RequeteSQL = RequeteSQL & "WHERE "
        End If
        If Filtre1 <> "" Then
            RequeteSQL = RequeteSQL & Filtre1
        End If
        If Filtre2 <> "" Then
            If Filtre1 <> "" Then RequeteSQL = RequeteSQL & " AND "
            RequeteSQL = RequeteSQL & Filtre2
        End If

        '--- fin de la requête ---
        RequeteSQL = RequeteSQL & "ORDER BY Premisses.NumPosteDepart, Premisses.NumPosteArrivee"
        'Debug.Print RequeteSQL

        With ADODCPremisses

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
                    Bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_121)
                End If
            End With

        End With

    End If
    'Debug.Print RequeteSQL

    '--- curseur de la souris ---
    SourisEnAttente False
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la grille de recherche
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheGrilleRecherche()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const HauteurPBCriteresRecherche As Integer = 1095
    
    '--- déclaration ---
    Dim HauteurGrilleRecherche As Long
    
    '--- affichage ---
    If RechercherSurGrille = False Then
        PBCriteresRecherche.Height = HauteurPBCriteresRecherche
        TDBGGrilleRecherche.Visible = False
        Me.Refresh
    Else
        PBCriteresRecherche.Height = PBBoutons.Top - PBCriteresRecherche.Top - MARGES.M_BORD_BAS - 8 * Screen.TwipsPerPixelY
        TDBGGrilleRecherche.Visible = True
    End If
    
    '--- hauteur de la grille de recherche ---
    HauteurGrilleRecherche = PBCriteresRecherche.Height - TDBGGrilleRecherche.Top - TDBGGrilleRecherche.Left - 5 * Screen.TwipsPerPixelY
    If HauteurGrilleRecherche > 0 Then
        TDBGGrilleRecherche.Height = HauteurGrilleRecherche
    End If
    
    '--- placer le focus ---
    If TDBGGrilleRecherche.Visible = True Then TDBGGrilleRecherche.SetFocus
    
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    If CBRechercherSurGrille.Visible = True Then CBRechercherSurGrille.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyF3: CBRechercherSurGrille_Click
        Case vbKeyF5:
        Case vbKeyF6:
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

Private Sub MEBEditionDetailsPremisses_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim TexteComplet As String, _
            TexteSansMasque As String

    If InterdireEvenements = False Then
    
        '--- affectation ---
        With MEBEditionDetailsPremisses
            TexteComplet = .Text
            TexteSansMasque = .ClipText
        End With
    
        '--- analyse en fonction de chaque colonne ---
        With MSHFGDetailsPremisses
                    
            Select Case .Col
                
                Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                    '--- code de l'action ---
                    InsertionDetail TexteSansMasque
                    GestionBoutons E_MODIFICATION_EN_COURS
                
                Case COLONNES_DETAILS_PREMISSES.C_PARAMETRE
                    '--- paramètres ---
                    TDetailsPremisses(.Row).Parametre = TexteSansMasque
                    GestionBoutons E_MODIFICATION_EN_COURS
                
                Case Else
    
            End Select
    
        End With

    End If

End Sub

Private Sub MEBEditionDetailsPremisses_GotFocus()
    On Error Resume Next
    SFocusTableDetailsPremisses.Visible = True
End Sub

Private Sub MEBEditionDetailsPremisses_KeyDown(KeyCode As Integer, Shift As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    With MSHFGDetailsPremisses

        '--- analyse en fonction de la touche ---
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

Private Sub MEBEditionDetailsPremisses_KeyPress(KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim TexteComplet As String, _
            TexteSansMasque As String

    With MSHFGDetailsPremisses

        '--- analyse de la touche ---
        Select Case KeyAscii

            Case vbKeyReturn
                '--- touche entrée ---
        
                '--- affectation ---
                With MEBEditionDetailsPremisses
                    TexteComplet = .Text
                    TexteSansMasque = .ClipText
                End With
                
                Select Case .Col

                    Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                        InsertionDetail TexteSansMasque
                        .Col = COLONNES_DETAILS_PREMISSES.C_PARAMETRE

                    Case COLONNES_DETAILS_PREMISSES.C_PARAMETRE
                        TDetailsPremisses(.Row).Parametre = TexteSansMasque
                        If .Row < .Rows - 1 Then .Row = .Row + 1
                        .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                        
                    Case Else

                End Select

                '--- mettre le focus sur le tableau ---
                .SetFocus
                KeyAscii = 0

            Case Else
                Select Case .Col
                    Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES, 10
                    Case COLONNES_DETAILS_PREMISSES.C_PARAMETRE: FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES, 10
                    Case Else
                End Select

        End Select

    End With

End Sub

Private Sub MEBEditionDetailsPremisses_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- focus ---
    SFocusTableDetailsPremisses.Visible = False
    
    '--- rendre le contrôle texte invisible ---
    MEBEditionDetailsPremisses.Visible = False

    '--- construction de la grille ---
    GestionDetailsPremisses GG_AFFICHAGE

End Sub

Private Sub MSHFGDetailsPremisses_DblClick()
    On Error Resume Next
    InterdireEvenements = True
    EditionDetailsPremisses vbKeySpace 'simule un espace
    InterdireEvenements = False
End Sub

Private Sub MSHFGDetailsPremisses_GotFocus()
    On Error Resume Next
    SFocusTableDetailsPremisses.Visible = True
End Sub

Private Sub MSHFGDetailsPremisses_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyDelete: EditionDetailsPremisses vbKeyBack 'simule un retour arrière (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGDetailsPremisses_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionDetailsPremisses KeyAscii  'envoi de la touche frappée
End Sub

Private Sub MSHFGDetailsPremisses_LeaveCell()
    On Error Resume Next
    MEBEditionDetailsPremisses.Visible = False
End Sub

Private Sub MSHFGDetailsPremisses_LostFocus()
    On Error Resume Next
    SFocusTableDetailsPremisses.Visible = False
End Sub

Private Sub MSHFGDetailsPremisses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- mémorisation de la ligne de départ ---
    With MSHFGDetailsPremisses
        If Button = vbKeyLButton And .MouseCol = COLONNES_DETAILS_PREMISSES.C_NUM_LIGNES Then
            LigneDepartDeplacement = .MouseRow
        End If
    End With

End Sub

Private Sub MSHFGDetailsPremisses_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- déclaration ---
    Dim TexteCellule As String
    Static MemTexteCellule As String
    
    '--- mémorisation de la ligne de départ ---
    With MSHFGDetailsPremisses
        
        '--- RAZ des variables de déplacement ---
        If Button <> vbKeyLButton Then
            LigneDepartDeplacement = 0
            LigneArriveeDeplacement = 0
        End If
        
        '--- affectation ---
        TexteCellule = .TextMatrix(.MouseRow, .MouseCol)
        
        If TexteCellule <> MemTexteCellule Then
        
            '--- gestion de la bulle ---
            Select Case .MouseCol
            
                Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                    '--- code de l'action ---
                    .ToolTipText = ""
                    If TexteCellule <> "" And .MouseRow > 0 Then
                        .ToolTipText = " Numéro de l'action = " & TDetailsPremisses(.MouseRow).NumAction & UN_ESPACE
                    End If

                Case Else
                    .ToolTipText = ""
        
            End Select
    
            '--- affectation ---
            MemTexteCellule = TexteCellule
    
        End If
    
    End With

End Sub

Private Sub MSHFGDetailsPremisses_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
        
    '--- mémorisation de la ligne d'arrivée ---
    With MSHFGDetailsPremisses
        If Button = vbKeyLButton And .MouseCol = COLONNES_DETAILS_PREMISSES.C_NUM_LIGNES Then
            LigneArriveeDeplacement = .MouseRow
            If LigneDepartDeplacement > 0 And _
               LigneArriveeDeplacement > 0 And _
               LigneDepartDeplacement <> LigneArriveeDeplacement Then
                    DeplacementLigne
            End If
        End If
    End With

End Sub

Private Sub MSHFGDetailsPremisses_Scroll()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rendre invisible le champ d'édition en cas de scrolling ---
    If MEBEditionDetailsPremisses.Visible = True Then
        MEBEditionDetailsPremisses.Visible = False
    End If

End Sub

Private Sub OBFormeGrille_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer                                                      'pour les boucles FOR...NEXT
    Dim NbrFractionnements As Integer                       'nombre de fractionnement
    
    '--- changement de la forme d'affichage ---
    SourisEnAttente True
    With TDBGGrilleRecherche
        Select Case Index
            
            Case 0
                '--- remettre en présentation normale ---
                .DataView = dbgNormalView               'présentation normale
                .Splits(0).AllowSizing = True               'autorise le fractionnement de la grille (petite rectangle noir en bas à gauche)
            
            Case 1
                '--- changement de la présentation ---
                NbrFractionnements = .Splits.Count
                If NbrFractionnements > 1 Then
                    For a = 2 To NbrFractionnements
                        .Splits.Remove 1                         'effacer le fractionnement 1 quelque soit le nombre de fractionnements
                    Next a
                End If
                .DataView = dbgInvertedView              'présentation inversée
            
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
    ADODCPremisses.Left = CBAnnuler.Left - MARGES.M_ENTRE_BOUTONS - ADODCPremisses.Width
    LRenseignements.Left = ADODCPremisses.Left
    CBActualiser.Left = ADODCPremisses.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
    CBSupprimer.Left = CBActualiser.Left - MARGES.M_ENTRE_BOUTONS - CBSupprimer.Width
    CBPremisseAutomatique.Left = CBSupprimer.Left - MARGES.M_ENTRE_BOUTONS - CBPremisseAutomatique.Width
    CBRegenerationComplete.Left = CBPremisseAutomatique.Left - MARGES.M_ENTRE_BOUTONS - CBRegenerationComplete.Width
    AfficheGrilleRecherche
    
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
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            ADODCPremisses.Enabled = True
            CBActualiser.Enabled = True
            PBCriteresRecherche.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
            ADODCPremisses.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
            ADODCPremisses.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
            ADODCPremisses.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            ADODCPremisses.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
            MarqueEnregistrement True
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBAnnuler.Enabled = True
            ADODCPremisses.Enabled = False
            CBActualiser.Enabled = False
            CBSupprimer.Enabled = False
            PBCriteresRecherche.Enabled = False

        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- après nouveau ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            PBCriteresRecherche.Enabled = False
            CBQuitter.Enabled = True
            CBValider.Enabled = True
            CBAnnuler.Enabled = True
            ADODCPremisses.Enabled = False
            CBActualiser.Enabled = False
            CBSupprimer.Enabled = False
            'Me.TBNomGamme.SetFocus
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            ADODCPremisses.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---
            CBQuitter.Enabled = True
            CBValider.Enabled = False
            CBAnnuler.Enabled = False
            CBActualiser.Enabled = True
            CBSupprimer.Enabled = True
            PBCriteresRecherche.Enabled = True
        
        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des détails des prémisses
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsPremisses(ByVal EtatSouhaite As GESTION_GRILLES)
    
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
    Dim PremisseDecodee As String
    Dim FicheVide As ImgDetailsPremisses, _
            TCopieDetailsPremisses(1 To NBR_LIGNES_DETAILS_PREMISSES) As ImgDetailsPremisses
    Dim TPremisseDecodee As Variant           'tableau de base contenant les actions après décodage

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des détails ---
            For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
                TDetailsPremisses(a) = FicheVide
            Next a

            '--- initialisation de la grille des détails ---
            With MSHFGDetailsPremisses

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_PREMISSES + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_PREMISSES + .FixedCols
                .RowHeight(0) = 400                    'épaisseur des titres
                .Row = 0

                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_PREMISSES.C_NUM_LIGNES
                .ColWidth(.Col) = 4 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                .ColWidth(.Col) = 14 * EPAISSEUR_CARACTERE: .Text = "Code"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_PREMISSES.C_PARAMETRE
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Paramètre"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_PREMISSES.C_LIBELLE_ACTION
                .ColWidth(.Col) = 50 * EPAISSEUR_CARACTERE: .Text = "Libellé de l'action"
                .ColAlignment(.Col) = flexAlignLeftCenter

                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
                
                    '--- N° de lignes ---
                    .Col = COLONNES_DETAILS_PREMISSES.C_NUM_LIGNES
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_PREMISSES.C_LIBELLE_ACTION
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.VERT_1, COULEURS.CYAN_1)
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
                TDetailsPremisses(a) = FicheVide
            Next a
            With MSHFGDetailsPremisses()
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- décodage ---
            PremisseDecodee = ADODCPremisses.Recordset.Fields("PremisseDecodee")
            If PremisseDecodee <> "" Then
                
                '--- construction du tableau ---
                TPremisseDecodee = Split(PremisseDecodee, SEPARATEUR_PREMISSES)
            
                '--- transfert des données dans le tableau ---
                PtrLigne = 1
                For a = LBound(TPremisseDecodee) To UBound(TPremisseDecodee)
                    
                    With TDetailsPremisses(PtrLigne)
                        
                        If TPremisseDecodee(a) <> "" Then
                                
                            '--- numéro de l'action ---
                            .NumAction = TPremisseDecodee(a)
                                
                            If .NumAction >= LBound(TActions()) And .NumAction <= UBound(TActions()) Then
                                                        
                                '--- remplir l'action ---
                                .CodeAction = TActions(.NumAction).CodeAction
                                .LibelleAction = TActions(.NumAction).LibelleAction
                                .ParametreOuiNon = TActions(.NumAction).ParametreOuiNon
                                
                                '--- contrôle sur le paramètre ---
                                If .ParametreOuiNon = False Then
                            
                                    '--- action sans paramètre ---
                                    .Parametre = ""
                                                        
                                Else
                            
                                    '--- action avec paramètre ---
                                    Inc a
                                    If a <= UBound(TPremisseDecodee) Then
                                        .Parametre = TPremisseDecodee(a)
                                    End If
                            
                                End If
                                    
                                '--- incrément de la ligne ---
                                Inc PtrLigne
                            
                            End If
                    
                        End If
                        
                    End With
                
                Next a

            End If

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---
            PtrLigne = 1
            For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
                If TDetailsPremisses(a).CodeAction <> "" Then
                    TCopieDetailsPremisses(PtrLigne) = TDetailsPremisses(a)
                    Inc PtrLigne
                End If
            Next a
            For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
                TDetailsPremisses(a) = TCopieDetailsPremisses(a)
            Next a

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGDetailsPremisses

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
                    .Row = a
                    If TDetailsPremisses(a).CodeAction = "" Then
                        TDetailsPremisses(a) = FicheVide
                        For b = 1 To NBR_COLONNES_DETAILS_PREMISSES
                            .Col = b
                            .Text = ""
                        Next b
                    Else
                        .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                        .Text = TDetailsPremisses(a).CodeAction
                        
                        .Col = COLONNES_DETAILS_PREMISSES.C_PARAMETRE
                        If TDetailsPremisses(a).ParametreOuiNon = True Then
                            If TDetailsPremisses(a).Parametre = "" Then
                                .Text = "X"
                            Else
                        
                                '--- paramètre par défaut ---
                                .Text = TDetailsPremisses(a).Parametre
                                    
                                '--- vérification du type de paramètre ---
                                Select Case TDetailsPremisses(a).CodeAction
                                    
                                    Case CODE_TEMPO
                                        '--- ne rien changer pour l'affichage ---
                                    
                                    Case CODE_TEMPO_EGOUTTAGE
                                        '--- temporisation d'égouttage ---
                                        .Text = "GAMME"                       'remplacement du texte par "GAMME"
                                                                                           'indiquant que le paramêtre se trouve
                                                                                           'dans la gamme
                                    Case CODE_TEMPO_STABILISATION
                                        '--- temporisation de stabilisation ---
                                        .Text = CStr(TEMPS_MINI_STABILISATION_AVEC_CHARGE) & " + GAMME"
                                        'remplacement du texte par le temps de stabilisation mini de la charge plus le texte
                                        ' " + GAMME" indiquant que le délai supplémentaire de stabilisation est donné au
                                        'chargement de la gamme
                                    
                                    Case Else
                                        '--- indiquer le poste quand le paramètre est un numéro de poste ---
                                        If IsNumeric(TDetailsPremisses(a).Parametre) = True Then
                                            If TDetailsPremisses(a).Parametre >= POSTES.P_CHGT_1 And TDetailsPremisses(a).Parametre <= DERNIER_POSTE Then
                                                .Text = .Text & " (" & TActions(TDetailsPremisses(a).Parametre).CodeAction & ")"
                                            End If
                                        End If
         
                                End Select
                            
                            End If
                        Else
                            .Text = ""
                        End If
                        
                        .Col = COLONNES_DETAILS_PREMISSES.C_LIBELLE_ACTION
                        .Text = TDetailsPremisses(a).LibelleAction
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
' Rôle      : Insertion d'un détail dans la grille des détails
' Entrées : CodeAction -> Code de l'action
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InsertionDetail(Optional ByVal CodeAction As Variant)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim ChaineDeRecherche As String
    Dim FicheVide As ImgDetailsPremisses

    '--- lancer la modification ---
    GestionBoutons E_MODIFICATION_EN_COURS

    If IsMissing(CodeAction) = True Then

        '--- les données viennent de la grille des codes ---
        For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
            With TDetailsPremisses(a)
                If .CodeAction = "" Then
                    
                    '--- affectation ---
                    .NumAction = ADODCActions.Recordset("NumAction").value
                    .CodeAction = TActions(.NumAction).CodeAction
                    .ParametreOuiNon = TActions(.NumAction).ParametreOuiNon
                    .LibelleAction = TActions(.NumAction).LibelleAction
        
                    With MSHFGDetailsPremisses
                        If .RowIsVisible(a) = False Then
                            .TopRow = a
                        End If
                        .Row = a
                        .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                    End With

                    Exit For

                End If
            End With
        Next a
        GestionDetailsPremisses GG_AFFICHAGE
        MSHFGDetailsPremisses.SetFocus

    Else

        '--- le numéro vient directement de la grille des détails ---
        With TDetailsPremisses(MSHFGDetailsPremisses.Row)

            '--- affectation ---
            ChaineDeRecherche = "CodeAction = '" & CodeAction & "'"

            '--- recherche du premier enregistrement ---
            ADODCActions.Recordset.MoveFirst
            ADODCActions.Recordset.Find ChaineDeRecherche

            '--- analyse après recherche ---
            If ADODCActions.Recordset.BOF = False And ADODCActions.Recordset.EOF = False Then
                .NumAction = ADODCActions.Recordset("NumAction").value
                .CodeAction = TActions(.NumAction).CodeAction
                .ParametreOuiNon = TActions(.NumAction).ParametreOuiNon
                .LibelleAction = TActions(.NumAction).LibelleAction
            Else
                TDetailsPremisses(MSHFGDetailsPremisses.Row) = FicheVide
            End If

        End With

    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Déplace une ligne dans la grille des détails
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DeplacementLigne()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            PtrLigne As Integer
    Dim TFicheDepart As ImgDetailsPremisses, _
            TCopieDetailsPremisses(1 To NBR_LIGNES_DETAILS_PREMISSES)    As ImgDetailsPremisses

    If LigneDepartDeplacement > 0 And LigneDepartDeplacement < NBR_LIGNES_DETAILS_PREMISSES And _
       LigneArriveeDeplacement > 0 And LigneArriveeDeplacement < NBR_LIGNES_DETAILS_PREMISSES And _
       LigneDepartDeplacement <> LigneArriveeDeplacement Then

        '--- affectation ---
        TFicheDepart = TDetailsPremisses(LigneDepartDeplacement)

        '--- suppression à la ligne de départ ---
        PtrLigne = 1
        For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
            If a <> LigneDepartDeplacement Then
                TCopieDetailsPremisses(PtrLigne) = TDetailsPremisses(a)
                Inc PtrLigne
            End If
        Next a

        '--- transfert dans le tableau ---
        For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
            TDetailsPremisses(a) = TCopieDetailsPremisses(a)
        Next a

        '--- fixation de l'arrivée en fonction du sens de déplacement ---
        If LigneArriveeDeplacement > LigneDepartDeplacement Then
            LigneArriveeDeplacement = Pred(LigneArriveeDeplacement)
        End If
        If LigneArriveeDeplacement < 1 Then LigneArriveeDeplacement = 1
        If LigneArriveeDeplacement > NBR_LIGNES_DETAILS_PREMISSES Then LigneArriveeDeplacement = NBR_LIGNES_DETAILS_PREMISSES

        '--- insertion à la ligne d'arrivée ---
        PtrLigne = 1
        For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
            If a = LigneArriveeDeplacement Then
                TCopieDetailsPremisses(PtrLigne) = TFicheDepart
                Inc PtrLigne
            End If
            If PtrLigne <= NBR_LIGNES_DETAILS_PREMISSES Then
                TCopieDetailsPremisses(PtrLigne) = TDetailsPremisses(a)
                Inc PtrLigne
            End If
            If PtrLigne >= NBR_LIGNES_DETAILS_PREMISSES Then Exit For
        Next a

        '--- transfert dans le tableau ---
        For a = 1 To NBR_LIGNES_DETAILS_PREMISSES
            TDetailsPremisses(a) = TCopieDetailsPremisses(a)
        Next a

        '--- reconstruction de la grille ---
        GestionDetailsPremisses GG_AFFICHAGE

        '--- gestion des boutons ---
        GestionBoutons E_MODIFICATION_EN_COURS
    
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture de la partie concernant le moteur d'inférence dans les prémisses
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LecturePartieIA()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim NumPosteDepart As Integer, _
            NumPosteArrivee As Integer
    Dim TempsCycleSecondes As Long
    
    '--- affectation ---
    If MemDernierBouton <> ETATS_BOUTONS.E_AVANT_NOUVEAU And _
       MemDernierBouton <> ETATS_BOUTONS.E_APRES_NOUVEAU Then
    
        '--- curseur de la souris ---
        SourisEnAttente True
        
        With ADODCPremisses.Recordset

            If Not .BOF And Not .EOF Then

                If .status = adRecOK Then
        
                    '--- affectation ---
                    NumPosteDepart = .Fields("NumPosteDepart")
                    NumPosteArrivee = .Fields("NumPosteArrivee")
    
                    '--- extraction du n° du pont donné par le moteur d'inférence et du temps de cycle par apprentissage ---
                    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
                       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then
                
                        '--- affichage du n° de pont IA ---
                        LNomPontIA.Caption = TPremisses(NumPosteDepart, NumPosteArrivee).NumPontIA

                        'ComboPont.ListIndex = 0
             
                        

                        
                        
                        '--- calcul du temps de cycle en secondes (temps théorique par apprentissage des temps de mouvements) ---
                        With TPremisses(NumPosteDepart, NumPosteArrivee)
                            
                            '--- calcul ---
                            If CalculTempsCyclePremisse(NumPosteDepart, NumPosteArrivee, TempsCycleSecondes) = OK Then
                                .TempsCycleSecondes = TempsCycleSecondes
                            Else
                                .TempsCycleSecondes = 0
                            End If
                            
                            '--- affichage ---
                            If .TempsCycleSecondes = 0 Then
                                LTempsCycleSecondes.Caption = ""
                            Else
                                LTempsCycleSecondes.Caption = .TempsCycleSecondes & " secondes " & "(" & CTemps2(.TempsCycleSecondes) & ")"
                            End If
                        
                        End With
                    
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
' Rôle      : Lecture des détails des prémisses
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureDetailsPremisses()

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- déclaration ---
    Dim a As Integer

    If MemDernierBouton <> ETATS_BOUTONS.E_AVANT_NOUVEAU And _
       MemDernierBouton <> ETATS_BOUTONS.E_APRES_NOUVEAU Then

        '--- curseur de la souris ---
        SourisEnAttente True

        '--- vidage de la grille ---
        GestionDetailsPremisses GG_VIDAGE
        GestionDetailsPremisses GG_AFFICHAGE
        
        With ADODCPremisses.Recordset

            If Not .BOF And Not .EOF Then

                If .status = adRecOK Then

                    If IsError(.Fields("PremisseDecodee")) = False Then
                        GestionDetailsPremisses GG_TRANSFERT_DONNEES
                        GestionDetailsPremisses GG_AFFICHAGE
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
' Rôle      : Permet l'édition des détails d'une prémisse
' Entrées : KeyAscii -> Code ASCII de la touche frappée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionDetailsPremisses(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- édition uniquement sur les bonnes colonnes ---
    Select Case MSHFGDetailsPremisses.Col

        Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION, _
                 COLONNES_DETAILS_PREMISSES.C_PARAMETRE

            With MEBEditionDetailsPremisses

                '--- affiche le contrôle texte au bon endroit (dans la cellule) ---
                .Move MSHFGDetailsPremisses.Left + MSHFGDetailsPremisses.CellLeft, _
                           MSHFGDetailsPremisses.Top + MSHFGDetailsPremisses.CellTop, _
                           MSHFGDetailsPremisses.CellWidth, _
                           MSHFGDetailsPremisses.CellHeight

                '--- paramètres de contrôle texte en fonction de la cellule ---
                .Mask = ""
                .Text = ""
                Select Case MSHFGDetailsPremisses.Col
                    Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION: .Mask = String(10, "A")
                    Case COLONNES_DETAILS_PREMISSES.C_PARAMETRE: .Mask = String(10, "A")
                    Case Else
                End Select

                '--- analyse du caractère qui a été tapé ---
                Select Case KeyAscii

                    Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                        '--- du code 0 à l'espace (sauf retour arrière, Entrée) cela signifie une modification du texte en cours ---
                        .SelText = MSHFGDetailsPremisses.Text
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .Visible = True
                        .SetFocus

                    Case vbKeyBack
                        '--- touche retour arrière ---
                        .SelText = ""
                        .Visible = True
                        .SetFocus
                        DoEvents
                        MEBEditionDetailsPremisses_Change

                    Case vbKeyReturn
                        '--- touche Entrée ---
                        With MSHFGDetailsPremisses
                            Select Case .Col
                                Case COLONNES_DETAILS_PREMISSES.C_CODE_ACTION: .Col = COLONNES_DETAILS_PREMISSES.C_PARAMETRE
                                Case COLONNES_DETAILS_PREMISSES.C_PARAMETRE
                                    If .Row < .Rows - 1 Then .Row = .Row + 1
                                    .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
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
' Entrées : NumPosteDepart -> Poste de départ
'                NumPosteArrivee -> Poste d'arrivée
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal NumPosteDepart As Integer, _
                                                    ByVal NumPosteArrivee As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- initialisation des champs / grilles ---
    GestionGrilleRecherche GG_INITIALISATION
    GestionGrilleRecherche GG_AFFICHAGE
    
    '--- lien avec les données
    LieDonneesRecherche NumPosteDepart, NumPosteArrivee
    LanceRechercheOuTri True
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---

    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    
    '--- images des fonds ---
    Me.Picture = ImgFondOrange2
    PBBoutons.Picture = ImgFondDesBoutons

    '--- divers sur ADO ---

    '--- divers sur les renseignements ---
    LRenseignements.BackColor = COULEURS.CYAN_0

    '--- divers sur la grille des actions ---
    With DGActions()
        .BackColor = COULEURS.JAUNE_0
        .ForeColor = COULEURS.BLEU_5
    End With

    '--- transfert des postes de départ et arrivée ---
    With CBNumPosteDepart
        .Clear
        .AddItem ("")
        .ItemData(.NewIndex) = 0
    End With
    With CBNumPosteArrivee
        .Clear
        .AddItem ("")
        .ItemData(.NewIndex) = 0
    End With
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a).DefinitionPoste
            CBNumPosteDepart.AddItem (.NomPoste & " - " & .LibellePoste)
            CBNumPosteDepart.ItemData(CBNumPosteDepart.NewIndex) = .NumPoste
            CBNumPosteArrivee.AddItem (.NomPoste & " - " & .LibellePoste)
            CBNumPosteArrivee.ItemData(CBNumPosteArrivee.NewIndex) = .NumPoste
        End With
    Next a
    
    '--- gestion des détails ---
    GestionDetailsPremisses GG_INITIALISATION

    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE
    
    
    With ComboPont
        .Clear
        .AddItem ("")
        .ItemData(.NewIndex) = 0
        .AddItem ("Pont 1")
        .ItemData(.NewIndex) = 1
        .AddItem ("Pont 2")
        .ItemData(.NewIndex) = 2
    End With
    
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : gére le copier coller spécial
'                ATTENTION - CETTE ROUTINE EST APPELEE PAR LA FONCTION COLLAGE SPECIALE
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionCopierCollerSpecial()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- analyse en fonction de la fenetre de copie ---
    Select Case NumFenetreEnCopie

        Case FENETRES.F_PREMISSES
            '--- prémisses vers prémisses (transfert de la prémisse décodée) ---
            With ADODCPremisses.Recordset
                If Not .BOF And Not .EOF Then
                    .Fields("PremisseDecodee") = CleDeCopie
                    GestionDetailsPremisses GG_VIDAGE
                    GestionDetailsPremisses GG_TRANSFERT_DONNEES
                    GestionDetailsPremisses GG_AFFICHAGE
                    GestionBoutons E_MODIFICATION_EN_COURS
                End If
            End With

        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    
    '--- réactualisation des prémisses ---
    ChargePremisses

    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFPremisses = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Construit les prémisses (codée et décodée)
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ConstruitPremisses(ByRef PremisseCodee As String, _
                                                     ByRef PremisseDecodee As String)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer

    '--- affectation ---
    PremisseCodee = ""
    PremisseDecodee = ""

    '--- construction des 2 prémisses ---
    For a = LBound(TDetailsPremisses()) To UBound(TDetailsPremisses())
        
        With TDetailsPremisses(a)
            
            If .NumAction = 0 And .CodeAction = "" Then
                
                '--- sortie directe si plus d'action ---
                Exit For
            
            Else
                
                '--- ajout du code de l'action dans la prémisse codée ---
                PremisseCodee = PremisseCodee & .CodeAction & SEPARATEUR_PREMISSES
                
                '--- ajout du numéro de l'action dans la prémisse décodée ---
                PremisseDecodee = PremisseDecodee & .NumAction & SEPARATEUR_PREMISSES
                
                '--- analyse du paramètre ---
                If .ParametreOuiNon = True Then
                    PremisseCodee = PremisseCodee & .Parametre & SEPARATEUR_PREMISSES
                    PremisseDecodee = PremisseDecodee & .Parametre & SEPARATEUR_PREMISSES
                End If
                
            End If
        
        End With
    
    Next a

    '--- élimination du dernier séparateur pour les deux prémisses ---
    If PremisseCodee <> "" Then
        PremisseCodee = Mid(PremisseCodee, 1, Pred(Len(PremisseCodee)))
    End If
    If PremisseDecodee <> "" Then
        PremisseDecodee = Mid(PremisseDecodee, 1, Pred(Len(PremisseDecodee)))
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
' Rôle      : Gestion de la grille de recherche
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionGrilleRecherche(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    
    '--- affectation ---

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGGrilleRecherche
                
                .Visible = False                                                            'rendre la grille invisible
                '.ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'autorise le fractionnement de la grille (petite rectangle noir en bas à gauche)
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.BLEU_5                      'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.JAUNE_0              'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.ORANGE_1  '.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
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
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGGrilleRecherche
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_POSTE_DEPART)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du poste de départ"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NOM_POSTE_DEPART)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du poste de départ"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_POSTE_ARRIVEE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du poste d'arrivée"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NOM_POSTE_ARRIVEE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du poste d'arrivée"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_PONT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du pont"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_PREMISSE_CODEE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Prémisse codée"
                    .Width = EPAISSEUR_CARACTERE * 84
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_PREMISSE_DECODEE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Prémisse décodée"
                    .Width = EPAISSEUR_CARACTERE * 60
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

Private Sub TDBGGrilleRecherche_Click()
    On Error Resume Next
    If Me.ActiveControl.Name <> TDBGGrilleRecherche.Name Then           'placer le focus sur la grille si nécessaire
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
    
    '--- déplacement du focus sur le bouton ---
    With SFocusGrilleRecherche
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height + Screen.TwipsPerPixelY
        .Width = ActiveControl.Width + Screen.TwipsPerPixelX
        .Visible = True
    End With

    '--- affichage de la barre de sélection ---
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
        'Case vbKeyF3, vbKeyReturn
        '    CBRechercherSurGrille_Click
        '    KeyCode = 0: Shift = 0
        'Case vbKeyF5
         '   CBLancerRecherche_Click
         '   KeyCode = 0: Shift = 0
        'Case vbKeyF6
        '    CBRaz_Click
        '    KeyCode = 0: Shift = 0
        Case vbKeyHome
            If Shift = vbCtrlMask Then
                ADODCPremisses.Recordset.MoveFirst
                KeyCode = 0: Shift = 0
            End If
        Case vbKeyEnd
            If Shift = vbCtrlMask Then
                ADODCPremisses.Recordset.MoveLast
                KeyCode = 0: Shift = 0
            End If
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown
        Case vbKeyTab
            If Shift = vbShiftMask Then
                'TBContenant.SetFocus
            Else
                'CBSupprimer.SetFocus
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

Private Sub TOBGestionGrilles_ButtonClick(ByVal Button As MSComctlLib.Button)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                                                    'pour les boucles FOR...NEXT
    Dim NumLigne As Integer                                     'numéro de ligne
    Dim FicheVide As ImgDetailsPremisses              'fiche vide à l'image des détails des prémisses
    
    '--- affectation ---

    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key

        Case "SupprimerLigne"
            '--- supprimer une ligne ---
            NumLigne = MSHFGDetailsPremisses.Row
            
            '--- suppression de la ligne ---
            If NumLigne > 0 And NumLigne <= NBR_LIGNES_DETAILS_PREMISSES Then
                If AppelFenetre(F_MESSAGE, _
                                        TITRE_MESSAGES, _
                                        MESSAGE_3 & CStr(NumLigne) & " ?", _
                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                        TYPES_BOUTONS.T_OUI_NON, _
                                        EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                    TDetailsPremisses(NumLigne) = FicheVide
                    GestionDetailsPremisses GG_COMPRESSION
                    GestionDetailsPremisses GG_AFFICHAGE
                    GestionBoutons E_MODIFICATION_EN_COURS
                End If
                MSHFGDetailsPremisses.SetFocus
            End If
        
        Case "CompacterGrille"
            '--- compacter la grille ---
            GestionDetailsPremisses GG_COMPRESSION
            GestionDetailsPremisses GG_AFFICHAGE
        
        Case "InsererLigne"
            '--- insérer ligne ---
            NumLigne = MSHFGDetailsPremisses.Row
        
            '--- suppression de la ligne ---
            If NumLigne > 0 And NumLigne <= NBR_LIGNES_DETAILS_PREMISSES Then
                For a = Pred(NBR_LIGNES_DETAILS_PREMISSES) To NumLigne Step -1
                    TDetailsPremisses(Succ(a)) = TDetailsPremisses(a)
                Next a
                TDetailsPremisses(NumLigne) = FicheVide
                GestionDetailsPremisses GG_AFFICHAGE
                With MSHFGDetailsPremisses
                    .Col = COLONNES_DETAILS_PREMISSES.C_CODE_ACTION
                    .SetFocus
                End With
            End If
    
        Case Else

    End Select

End Sub

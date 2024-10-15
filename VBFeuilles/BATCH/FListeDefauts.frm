VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FListeDefauts 
   Appearance      =   0  'Flat
   Caption         =   "LISTE DES DEFAUTS"
   ClientHeight    =   7095
   ClientLeft      =   1440
   ClientTop       =   3915
   ClientWidth     =   10215
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
   Icon            =   "FListeDefauts.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   10215
   WindowState     =   2  'Maximized
   Begin TrueOleDBGrid80.TDBDropDown TDBDDIntervenants 
      Bindings        =   "FListeDefauts.frx":014A
      Height          =   3345
      Left            =   7680
      TabIndex        =   9
      Top             =   960
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   5900
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "N?"
      Columns(0).DataField=   "NumIntervenant"
      Columns(0).DataWidth=   6
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NomPrenomIntervenant"
      Columns(1).DataField=   "NomPrenomIntervenant"
      Columns(1).DataWidth=   41
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "NumTelephone"
      Columns(2).DataField=   "NumTelephone"
      Columns(2).DataWidth=   20
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).ExtendRightColumn=   -1  'True
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowRowSizing=   0   'False
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=741"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=609"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=7911"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=7779"
      Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2778"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits.Count    =   1
      AllowRowSizing  =   0   'False
      Appearance      =   1
      BorderStyle     =   1
      ColumnHeaders   =   -1  'True
      DataMode        =   0
      DefColWidth     =   0
      Enabled         =   -1  'True
      HeadLines       =   1
      RowDividerStyle =   2
      LayoutName      =   ""
      LayoutFileName  =   ""
      LayoutURL       =   ""
      EmptyRows       =   0   'False
      ListField       =   "NomPrenomIntervenant"
      DataField       =   "NumIntervenant"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      RowTracking     =   -1  'True
      DeadAreaBackColor=   12632319
      ValueTranslate  =   -1  'True
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HFFFFC0&,.fgcolor=&H0&,.bold=-1"
      _StyleDefs(7)   =   ":id=1,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&HFFFF00&,.fgcolor=&H0&"
      _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000F&"
      _StyleDefs(14)  =   ":id=3,.fgcolor=&H80C0FF&"
      _StyleDefs(15)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(16)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(17)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(18)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(19)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(20)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(21)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(22)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(23)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(24)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(25)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(47)  =   "Named:id=33:Normal"
      _StyleDefs(48)  =   ":id=33,.parent=0"
      _StyleDefs(49)  =   "Named:id=34:Heading"
      _StyleDefs(50)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   ":id=34,.wraptext=-1"
      _StyleDefs(52)  =   "Named:id=35:Footing"
      _StyleDefs(53)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(54)  =   "Named:id=36:Selected"
      _StyleDefs(55)  =   ":id=36,.parent=33,.bgcolor=&HFF0000&"
      _StyleDefs(56)  =   "Named:id=37:Caption"
      _StyleDefs(57)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(58)  =   "Named:id=38:HighlightRow"
      _StyleDefs(59)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(60)  =   "Named:id=39:EvenRow"
      _StyleDefs(61)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(62)  =   "Named:id=40:OddRow"
      _StyleDefs(63)  =   ":id=40,.parent=33"
      _StyleDefs(64)  =   "Named:id=41:RecordSelector"
      _StyleDefs(65)  =   ":id=41,.parent=34"
      _StyleDefs(66)  =   "Named:id=42:FilterBar"
      _StyleDefs(67)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid80.TDBGrid TDBGListeDefauts 
      Align           =   1  'Align Top
      Bindings        =   "FListeDefauts.frx":016A
      Height          =   4515
      Left            =   0
      TabIndex        =   8
      Top             =   375
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7964
      _LayoutType     =   1
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NumDefaut"
      Columns(0).DataField=   "NumDefaut"
      Columns(0).DataWidth=   6
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "SignalerOuiNon"
      Columns(1).DataField=   "SignalerOuiNon"
      Columns(1).DataWidth=   6
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "GyrophareOuiNon"
      Columns(2).DataField=   "GyrophareOuiNon"
      Columns(2).DataWidth=   6
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "KlaxonOuiNon"
      Columns(3).DataField=   "KlaxonOuiNon"
      Columns(3).DataWidth=   6
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "MessageVocalOuiNon"
      Columns(4).DataField=   "MessageVocalOuiNon"
      Columns(4).DataWidth=   6
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "AfficheurOuiNon"
      Columns(5).DataField=   "AfficheurOuiNon"
      Columns(5).DataWidth=   6
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "InformationsAPI"
      Columns(6).DataField=   "InformationsAPI"
      Columns(6).DataWidth=   30
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "LibelleDefaut"
      Columns(7).DataField=   "LibelleDefaut"
      Columns(7).DataWidth=   100
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "LibelleDefautAfficheur"
      Columns(8).DataField=   "LibelleDefautAfficheur"
      Columns(8).DataWidth=   100
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "NumIntervenant1"
      Columns(9).DataField=   "NumIntervenant1"
      Columns(9).DataWidth=   6
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "NumIntervenant2"
      Columns(10).DataField=   "NumIntervenant2"
      Columns(10).DataWidth=   6
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "NumIntervenant3"
      Columns(11).DataField=   "NumIntervenant3"
      Columns(11).DataWidth=   6
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "NumIntervenant4"
      Columns(12).DataField=   "NumIntervenant4"
      Columns(12).DataWidth=   6
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "NumIntervenant5"
      Columns(13).DataField=   "NumIntervenant5"
      Columns(13).DataWidth=   6
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   14
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0)._GSX_SAVERECORDSELECTORS=   0
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=14"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2355"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2223"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(0)._AlignLeft=0"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3228"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3096"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(1)._AlignLeft=0"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=3625"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3493"
      Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(15)=   "Column(2)._AlignLeft=0"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2910"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2778"
      Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(20)=   "Column(3)._AlignLeft=0"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=4445"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=4313"
      Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(25)=   "Column(4)._AlignLeft=0"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3149"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(30)=   "Column(5)._AlignLeft=0"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=4366"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=4233"
      Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(35)=   "Column(7).Width=4366"
      Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=4233"
      Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(39)=   "Column(8).Width=4366"
      Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=4233"
      Splits(0)._ColumnProps(42)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(43)=   "Column(9).Width=3387"
      Splits(0)._ColumnProps(44)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(45)=   "Column(9)._WidthInPix=3254"
      Splits(0)._ColumnProps(46)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(47)=   "Column(9)._AlignLeft=0"
      Splits(0)._ColumnProps(48)=   "Column(10).Width=3387"
      Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=3254"
      Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(52)=   "Column(10)._AlignLeft=0"
      Splits(0)._ColumnProps(53)=   "Column(11).Width=3387"
      Splits(0)._ColumnProps(54)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(11)._WidthInPix=3254"
      Splits(0)._ColumnProps(56)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(57)=   "Column(11)._AlignLeft=0"
      Splits(0)._ColumnProps(58)=   "Column(12).Width=3387"
      Splits(0)._ColumnProps(59)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(12)._WidthInPix=3254"
      Splits(0)._ColumnProps(61)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(62)=   "Column(12)._AlignLeft=0"
      Splits(0)._ColumnProps(63)=   "Column(13).Width=3387"
      Splits(0)._ColumnProps(64)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(65)=   "Column(13)._WidthInPix=3254"
      Splits(0)._ColumnProps(66)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(67)=   "Column(13)._AlignLeft=0"
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=71,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=80,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=72,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=73,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=74,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=76,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=75,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=77,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=78,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=79,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=81,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=82,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=16,.parent=71"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=72"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=73"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=75"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=20,.parent=71"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=72"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=73"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=75"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=24,.parent=71"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=21,.parent=72"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=22,.parent=73"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=23,.parent=75"
      _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=28,.parent=71"
      _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=72"
      _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=73"
      _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=75"
      _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=32,.parent=71"
      _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=72"
      _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=73"
      _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=75"
      _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=46,.parent=71"
      _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=72"
      _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=73"
      _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=75"
      _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=50,.parent=71"
      _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=72"
      _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=73"
      _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=75"
      _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=54,.parent=71"
      _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=72"
      _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=73"
      _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=75"
      _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=58,.parent=71"
      _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=72"
      _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=73"
      _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=75"
      _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=62,.parent=71"
      _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=72"
      _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=73"
      _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=75"
      _StyleDefs(72)  =   "Splits(0).Columns(10).Style:id=66,.parent=71"
      _StyleDefs(73)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=72"
      _StyleDefs(74)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=73"
      _StyleDefs(75)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=75"
      _StyleDefs(76)  =   "Splits(0).Columns(11).Style:id=70,.parent=71"
      _StyleDefs(77)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=72"
      _StyleDefs(78)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=73"
      _StyleDefs(79)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=75"
      _StyleDefs(80)  =   "Splits(0).Columns(12).Style:id=86,.parent=71"
      _StyleDefs(81)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=72"
      _StyleDefs(82)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=73"
      _StyleDefs(83)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=75"
      _StyleDefs(84)  =   "Splits(0).Columns(13).Style:id=90,.parent=71"
      _StyleDefs(85)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=72"
      _StyleDefs(86)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=73"
      _StyleDefs(87)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=75"
      _StyleDefs(88)  =   "Named:id=33:Normal"
      _StyleDefs(89)  =   ":id=33,.parent=0"
      _StyleDefs(90)  =   "Named:id=34:Heading"
      _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   ":id=34,.wraptext=-1"
      _StyleDefs(93)  =   "Named:id=35:Footing"
      _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(95)  =   "Named:id=36:Selected"
      _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=37:Caption"
      _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(99)  =   "Named:id=38:HighlightRow"
      _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(101) =   "Named:id=39:EvenRow"
      _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(103) =   "Named:id=40:OddRow"
      _StyleDefs(104) =   ":id=40,.parent=33"
      _StyleDefs(105) =   "Named:id=41:RecordSelector"
      _StyleDefs(106) =   ":id=41,.parent=34"
      _StyleDefs(107) =   "Named:id=42:FilterBar"
      _StyleDefs(108) =   ":id=42,.parent=33"
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Picture         =   "FListeDefauts.frx":018A
      ScaleHeight     =   315
      ScaleWidth      =   10155
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   10215
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "LISTE DES DEFAUTS"
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
         Left            =   120
         TabIndex        =   5
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc ADODCListeDefauts 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5670
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   16777215
      ForeColor       =   0
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSOLEDBSQL18;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
      OLEDBString     =   "Provider=MSOLEDBSQL18;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM ListeDefauts ORDER BY NumDefaut"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   10155
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6000
      Width           =   10215
      Begin MSComctlLib.ImageList ILGrillesDonnees 
         Left            =   120
         Top             =   480
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
               Picture         =   "FListeDefauts.frx":24ACC
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":24CD8
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":24EE4
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":250F0
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":252FC
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":25508
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":25714
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":25920
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":25B2C
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":25D38
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":25F44
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26150
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":2635C
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26568
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26774
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26980
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26B8C
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26D90
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":26F94
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":27198
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":2739C
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":275A0
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":277A4
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":279A8
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":27BAC
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FListeDefauts.frx":27DB0
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CBPasSignalisationDefauts 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAS de signalisation"
         DownPicture     =   "FListeDefauts.frx":27FB4
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
         Left            =   8400
         MaskColor       =   &H00FF00FF&
         Picture         =   "FListeDefauts.frx":29556
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Pas de signalisation des d?fauts du chargement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton CBSignalerTousDefauts 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Signaler tous les d?fauts"
         DownPicture     =   "FListeDefauts.frx":2AAF8
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
         Left            =   5280
         MaskColor       =   &H00FF00FF&
         Picture         =   "FListeDefauts.frx":2BE5A
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Signale tous les d?fauts du chargement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton CBTra?abiliteAlarmes 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tra?abilit? des alarmes"
         DownPicture     =   "FListeDefauts.frx":2D1BC
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
         Left            =   2160
         MaskColor       =   &H00FF00FF&
         Picture         =   "FListeDefauts.frx":2DF7E
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   " Tra?abilit? totale des alarmes du chargement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "F10 = ACTUALISE&R"
         DownPicture     =   "FListeDefauts.frx":2ED40
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
         Left            =   11580
         MaskColor       =   &H00FF00FF&
         Picture         =   "FListeDefauts.frx":2F442
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Actualiser les donn?es "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "Echap=&QUITTER"
         DownPicture     =   "FListeDefauts.frx":2FB44
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
         Left            =   14700
         MaskColor       =   &H00FF00FF&
         Picture         =   "FListeDefauts.frx":30246
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Quitter cette fen?tre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   345
         Left            =   120
         Top             =   60
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc ADODCIntervenants 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   5340
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   16777215
      ForeColor       =   0
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSOLEDBSQL18;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
      OLEDBString     =   "Provider=MSOLEDBSQL18;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   $"FListeDefauts.frx":30948
      Caption         =   "ADODCIntervenants"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FListeDefauts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R?le                    : Fen?tre affichant la liste de l'ensemble des d?fauts possibles sur la ligne
' Nom                    : FListeDefauts.frm
' Date de cr?ation : 27/10/2010
' D?tails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d?clarations obligatoires ---
Option Explicit

'--- options g?n?rales ---
Option Base 1
DefVar A-Z

'--- constantes priv?es ---
Private Const EPAISSEUR_LIGNE As Integer = 230
Private Const TITRE_FENETRE As String = "LISTE DES DEFAUTS"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- ?num?rations priv?es ---
Private Enum COLONNES_DETAILS_LISTE_DEFAUTS
    C_NUM_DEFAUT = 0
    C_SIGNALER_OUI_NON = 1
    C_GYROPHARE_OUI_NON = 2
    C_KLAXON_OUI_NON = 3
    C_MESSAGE_VOCAL_OUI_NON = 4
    C_AFFICHEUR_OUI_NON = 5
    C_INFORMATIONS_API = 6
    C_LIBELLE_DEFAUT = 7
    C_LIBELLE_DEFAUT_AFFICHEUR = 8
    C_NUM_INTERVEANT_1 = 9
    C_NUM_INTERVEANT_2 = 10
    C_NUM_INTERVEANT_3 = 11
    C_NUM_INTERVEANT_4 = 12
    C_NUM_INTERVEANT_5 = 13
End Enum

Private Enum COLONNES_DETAILS_INTERVENANTS
    C_NUM_INTERVENANT = 0
    C_NOM_PRENOM_INTERVENANT = 1
    C_NUM_TELEPHONE = 2
End Enum

'--- variables priv?es ---
Private PremiereActivation As Boolean

'--- tableaux priv?s ---

'--- variables publiques ---
Public NumFenetre As Long                             'num?ro de la fen?tre lorsqu'elle devient active

Private Sub CBActualiser_Click()
    On Error Resume Next
    ADODCListeDefauts.Refresh
    ChargeDefauts
End Sub

Private Sub CBActualiser_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d?placement du focus sur le bouton ---
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

Private Sub CBPasSignalisationDefauts_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- signalisation d'aucun d?faut ---
    SignalerTousLesDefauts False
    
    '--- actualisation des donn?es ---
    CBActualiser_Click

End Sub

Private Sub CBPasSignalisationDefauts_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d?placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBPasSignalisationDefauts_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
    DechargeFenetre
End Sub

Private Sub CBQuitter_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d?placement du focus sur le bouton ---
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

Private Sub CBSignalerTousDefauts_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- signalisation de tous les d?fauts ---
    SignalerTousLesDefauts True
    
    '--- actualisation des donn?es ---
    CBActualiser_Click

End Sub

Private Sub CBSignalerTousDefauts_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d?placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBSignalerTousDefauts_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBTra?abiliteAlarmes_Click()
    On Error Resume Next
    AppelFenetre F_TRACABILITE_ALARMES
End Sub

Private Sub CBTra?abiliteAlarmes_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d?placement du focus sur le bouton ---
    With SFocus
        .Left = ActiveControl.Left
        .Top = ActiveControl.Top
        .Height = ActiveControl.Height
        .Width = ActiveControl.Width
        .Visible = True
    End With

End Sub

Private Sub CBTra?abiliteAlarmes_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fen?tre principale ---
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
    
    '--- gestion des touches relatifs ? cette Fenetre ---
    GestionTouches KeyCode, Shift
    
End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- dimensions de la grille ---
    TDBGListeDefauts.Height = Abs(Me.ScaleHeight - PBRenseignementsFenetre.Height - Me.ADODCListeDefauts.Height - PBBoutons.Height)

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
    
    '--- d?claration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBActualiser.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
    CBPasSignalisationDefauts.Left = CBActualiser.Left - MARGES.M_ENTRE_BOUTONS - CBPasSignalisationDefauts.Width
    CBSignalerTousDefauts.Left = CBPasSignalisationDefauts.Left - MARGES.M_ENTRE_BOUTONS - CBSignalerTousDefauts.Width
    CBTra?abiliteAlarmes.Left = CBSignalerTousDefauts.Left - MARGES.M_ENTRE_BOUTONS - CBTra?abiliteAlarmes.Width
    
    '--- recalcul du focus apr?s d?placement ---
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
' R?le      : D?charge la fen?tre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False

    '--- mise ? jour de la grille pour ?viter le message d'erreurs (data type mismatch during field update) ---
    TDBGListeDefauts.Update
    
    '--- r?actualisation des d?fauts ---
    ChargeDefauts
    
    '--- curseur souris par d?faut ---
    SourisEnAttente False
    
    '--- d?chargement de la fen?tre ---
    Me.Visible = False
    Unload Me
    Set OccFListeDefauts = Nothing

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R?le      : Change le curseur de la souris en fonction de l'attente
' Entr?es : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par d?faut
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
' R?le      : Initialise la fen?tre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d?claration ---

    '--- affectation ---
  
    '--- divers sur la fen?tre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .WindowState = vbMaximized
    End With
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- renseignements de la fen?tre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)

    '--- initialisation des grilles ---
    GestionListeDefauts GG_INITIALISATION
    GestionIntervenants GG_INITIALISATION
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R?le      : Gestion de la liste des intervenants
' Entr?es : EtatSouhaite -> Fonction de l'?num?ration GESTION_GRILLES
' Retours : "" indique aucun incident sinon le num?ro de l'erreur est renvoy?
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionIntervenants(ByVal EtatSouhaite As GESTION_GRILLES) As String

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv?es ---
    
    '--- d?claration ---

    '--- affectation ---
    GestionIntervenants = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBDDIntervenants
                
                .Visible = False                                                            'rendre la grille invisible
            
                .HeadLines = 2                                                             'nombre de ligne des ent?tes
                .HeadBackColor = COULEURS.BLEU_4                      'couleur de fond des ent?tes
                .HeadForeColor = COULEURS.JAUNE_3                     'couleur de plan des ent?tes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilis?e
                
                .AlternatingRowStyle = False                                       'pas de lignes en alternance
                .BackColor = COULEURS.BLANC                                 'couleur de fond
                .ForeColor = COULEURS.NOIR                                    'couleur de premier plan
                
                With .Styles(5)                                                             'couleur du curseur de s?lection
                    .BackColor = COULEURS.VERT_5
                    .ForeColor = COULEURS.BLANC
                End With
                
                With .HeadFont
                    .Name = "Arial"
                    .Bold = True                                                              'caract?res gras
                End With
                
                With .Font
                    .Name = "MS Sans serif"
                    .Bold = True                                                              'caract?res gras
                End With
                
                .RowHeight = 0                                                              '?paisseur des lignes
                .RowHeight = .RowHeight * 1.05
                
                .AllowColSelect = False                                                'interdire la s?lection des colonnes
                .AllowColMove = False                                                 'interdire le d?placement des colonnes s?lectionn?es
                .AllowRowSizing = False                                              'interdire la modification de l'?paisseur des lignes
                
                With .Columns(COLONNES_DETAILS_INTERVENANTS.C_NUM_INTERVENANT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N?"
                    .Width = EPAISSEUR_CARACTERE * 4
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_DETAILS_INTERVENANTS.C_NOM_PRENOM_INTERVENANT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom et pr?nom"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_INTERVENANTS.C_NUM_TELEPHONE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "T?l?phone"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionIntervenants = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R?le      : Gestion de la liste des d?fauts
' Entr?es : EtatSouhaite -> Fonction de l'?num?ration GESTION_GRILLES
' Retours : "" indique aucun incident sinon le num?ro de l'erreur est renvoy?
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionListeDefauts(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv?es ---
    
    '--- d?claration ---

    '--- affectation ---
    GestionListeDefauts = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGListeDefauts
                
                .Visible = False                                                            'rendre la grille invisible
                
                .Splits(0).AllowSizing = True                                        'autorise le fractionnement de la grille (petite rectangle noir en bas ? gauche)
            
                .HeadLines = 3                                                             'nombre de ligne des ent?tes
                .HeadBackColor = COULEURS.ROUGE_3                   'couleur de fond des ent?tes
                .HeadForeColor = COULEURS.JAUNE_3                     'couleur de plan des ent?tes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilis?e
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.CYAN_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                .ForeColor = COULEURS.BLEU_4                                'couleurs des donn?es
                
                .HeadFont.Name = "Arial"
                With .Font
                    .Name = "MS Sans serif"
                    .Bold = True                                                              'caract?res gras
                End With
                
                .RowHeight = 0                                                              '?paisseur des lignes
                .RowHeight = .RowHeight * 1.05
                
                .RecordSelectors = True                                                'affichage du s?lecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 '?paisseur du s?lecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du s?lecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du s?lecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = Me.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = Me.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = Me.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = False                                                    'interdire un nouvel enregistrement
                .AllowUpdate = True                                                       'autoriser la modification des donn?es
                
                .AllowDelete = False                                                      'interdire la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                 'interdire la s?lection des colonnes
                .AllowColMove = False                                                  'interdire le d?placement des colonnes s?lectionn?es
                
                .AllowRowSelect = True                                                  'autoriser la s?lection des lignes
                .AllowRowSizing = True                                                  'autoriser la modification de l'?paisseur des lignes
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_NUM_DEFAUT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N? du d?faut"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_SIGNALER_OUI_NON)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Signaler le d?faut (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_GYROPHARE_OUI_NON)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Avec gyrophare (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_KLAXON_OUI_NON)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Avec klaxon (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_MESSAGE_VOCAL_OUI_NON)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Message vocale (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_AFFICHEUR_OUI_NON)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Avec afficheur (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_INFORMATIONS_API)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Informations automate"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_LIBELLE_DEFAUT)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libell? du d?faut"
                    .Width = EPAISSEUR_CARACTERE * 60
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_LIBELLE_DEFAUT_AFFICHEUR)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libell? pour l'afficheur"
                    .Width = EPAISSEUR_CARACTERE * 60
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_NUM_INTERVEANT_1)
                    .Locked = False
                    .ValueItems.Presentation = dbgComboBox
                    .Caption = "Intervenant 1"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_NUM_INTERVEANT_2)
                    .Locked = False
                    .ValueItems.Presentation = dbgComboBox
                    .Caption = "Intervenant 2"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_NUM_INTERVEANT_3)
                    .Locked = False
                    .ValueItems.Presentation = dbgComboBox
                    .Caption = "Intervenant 3"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_NUM_INTERVEANT_4)
                    .Locked = False
                    .ValueItems.Presentation = dbgComboBox
                    .Caption = "Intervenant 4"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_DETAILS_LISTE_DEFAUTS.C_NUM_INTERVEANT_5)
                    .Locked = False
                    .ValueItems.Presentation = dbgComboBox
                    .Caption = "Intervenant 5"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
           
                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionListeDefauts = CStr(Err.Number)

End Function

Private Sub PBRenseignementsFenetre_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d?claration ---
    
    '--- calculs des emplacements ---
    With PBRenseignementsFenetre
        LRenseignementsFenetre.Left = .ScaleLeft
        LRenseignementsFenetre.Top = .ScaleTop + 30
        LRenseignementsFenetre.Width = .ScaleWidth
        LRenseignementsFenetre.Height = .ScaleHeight
    End With

End Sub

Private Sub TDBDDIntervenants_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = 0                'annulation de l'affichage des messages d'erreurs
End Sub

Private Sub TDBGListeDefauts_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = 0                'annulation de l'affichage des messages d'erreurs
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R?le      : G?re l'appui des touches du clavier
' Entr?es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionTouches(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- action en fonction des touches ---
    Select Case KeyCode
        
        Case vbKeyF10
            '--- touche F10 (actualiser) ---
            If CBActualiser.Enabled = True Then
                CBActualiser_Click
            End If
            KeyCode = 0
    
        Case Else
    End Select
    
End Sub

' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R?le      : Signaler ou pas dans la base de donn?es la totalit? des d?fauts
' Entr?es : AvecSignalisation ->  TRUE = Signale tous les d?fauts
'                                                  FALSE = Signale ancun d?faut
' Retours :
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SignalerTousLesDefauts(ByVal AvecSignalisation As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d?claration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- ouverture de la connexion ? la base de donn?es d'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- affectation de la requ?te ---
    Requete = "UPDATE ListeDefauts SET SignalerOuiNon = " & IIf(AvecSignalisation, -1, 0) & " WHERE (LibelleDefaut <> '')"

    '--- lancement de la requ?te ---
    Enregistrement.Open Requete, ConnexionBDAnodisationSQL, adOpenKeyset, adLockOptimistic
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Sub



VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FTraçabiliteAlarmes 
   BackColor       =   &H00C0C0C0&
   Caption         =   "TRACABILITE DES ALARMES"
   ClientHeight    =   10845
   ClientLeft      =   795
   ClientTop       =   3555
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FTraçabiliteAlarmes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10845
   ScaleWidth      =   13680
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   7.5
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Picture         =   "FTraçabiliteAlarmes.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   13620
      TabIndex        =   13
      Top             =   0
      Width           =   13680
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "TRACABILITE DES ALARMES"
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
         TabIndex        =   14
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
         Name            =   "Marlett"
         Size            =   7.5
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   13620
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   9750
      Width           =   13680
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   7.5
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   19
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FTraçabiliteAlarmes.frx":24A8C
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
            Picture         =   "FTraçabiliteAlarmes.frx":24C36
            Style           =   1  'Graphical
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   0
            Width           =   255
         End
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   20
            Top             =   720
            Width           =   915
         End
      End
      Begin MSComctlLib.ImageList ILGrillesDonnees 
         Left            =   1380
         Top             =   420
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
               Picture         =   "FTraçabiliteAlarmes.frx":24DE0
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":24FEC
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":251F8
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":25404
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":25610
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":2581C
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":25A28
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":25C34
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":25E40
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":2604C
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":26258
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":26464
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":26670
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":2687C
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":26A88
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":26C94
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":26EA0
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":270A4
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":272A8
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":274AC
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":276B0
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":278B4
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":27AB8
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":27CBC
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":27EC0
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteAlarmes.frx":280C4
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc ADODCTraçabiliteAlarmes 
         Height          =   390
         Left            =   15240
         Top             =   540
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   688
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   2
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   1000
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
         RecordSource    =   $"FTraçabiliteAlarmes.frx":282C8
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
      Begin VB.CommandButton CBSupprimer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Supprimer la totalité de la traçabilité"
         DownPicture     =   "FTraçabiliteAlarmes.frx":283DC
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
         Left            =   7380
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTraçabiliteAlarmes.frx":28ADE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Supprimer la totalité de la traçabilité des alarmes "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   4275
      End
      Begin VB.CommandButton CBQuitter 
         Cancel          =   -1  'True
         Caption         =   "Echap=&QUITTER"
         DownPicture     =   "FTraçabiliteAlarmes.frx":291E0
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
         Left            =   17280
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTraçabiliteAlarmes.frx":298E2
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "F10 = ACTUALISE&R"
         DownPicture     =   "FTraçabiliteAlarmes.frx":29FE4
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
         Left            =   11940
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTraçabiliteAlarmes.frx":2A6E6
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   " Actualiser les données "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3015
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   225
         Left            =   1380
         Top             =   120
         Visible         =   0   'False
         Width           =   300
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
         Height          =   285
         Left            =   15240
         TabIndex        =   12
         Top             =   180
         Width           =   1695
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   7.5
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   13995
      Index           =   0
      Left            =   0
      ScaleHeight     =   13995
      ScaleWidth      =   13680
      TabIndex        =   7
      Top             =   375
      Width           =   13680
      Begin VB.PictureBox PBDeplacementFenetre 
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   7.5
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   13635
         Index           =   1
         Left            =   0
         ScaleHeight     =   13575
         ScaleWidth      =   28620
         TabIndex        =   8
         Top             =   -15
         Width           =   28680
         Begin VB.Frame FGrilleTraçabiliteAlarmes 
            Caption         =   " Grille de la traçabilité des alarmes "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   11415
            Left            =   180
            TabIndex        =   16
            Top             =   1140
            Width           =   28335
            Begin TrueOleDBGrid80.TDBGrid TDBGTraçabiliteAlarmes 
               Bindings        =   "FTraçabiliteAlarmes.frx":2ADE8
               Height          =   10665
               Left            =   300
               TabIndex        =   18
               Top             =   420
               Width           =   27705
               _ExtentX        =   48869
               _ExtentY        =   18812
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "NumDefaut"
               Columns(0).DataField=   "NumDefaut"
               Columns(0).DataWidth=   11
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "DateDetectionDefaut"
               Columns(1).DataField=   "DateDetectionDefaut"
               Columns(1).DataWidth=   19
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "DateCorrectionDefaut"
               Columns(2).DataField=   "DateCorrectionDefaut"
               Columns(2).DataWidth=   19
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "LibelleDefaut"
               Columns(3).DataField=   "LibelleDefaut"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "ComplementDefaut"
               Columns(4).DataField=   "ComplementDefaut"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   5
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=5"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=3307"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3201"
               Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(5)=   "Column(0)._AlignLeft=0"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=5636"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5530"
               Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(10)=   "Column(1)._AlignLeft=0"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=5821"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5715"
               Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(15)=   "Column(2)._AlignLeft=0"
               Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
               Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
               Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(20)=   "Column(4).Width=3281"
               Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=3175"
               Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   0
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Microsoft Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Microsoft Sans Serif"
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
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=-1,.fontsize=750,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=2"
               _StyleDefs(5)   =   ":id=0,.fontname=Marlett"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=975,.italic=0"
               _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=Microsoft Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
               _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
               _StyleDefs(12)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(13)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(14)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(15)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(16)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(17)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(18)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(19)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(20)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(21)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(22)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(23)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(24)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(25)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(26)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(27)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(28)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(29)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(30)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(31)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(32)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(33)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(34)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
               _StyleDefs(35)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
               _StyleDefs(36)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
               _StyleDefs(37)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
               _StyleDefs(38)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
               _StyleDefs(39)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
               _StyleDefs(40)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
               _StyleDefs(41)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
               _StyleDefs(42)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
               _StyleDefs(43)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
               _StyleDefs(44)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
               _StyleDefs(45)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
               _StyleDefs(46)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
               _StyleDefs(47)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
               _StyleDefs(48)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
               _StyleDefs(49)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
               _StyleDefs(50)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
               _StyleDefs(51)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
               _StyleDefs(52)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
               _StyleDefs(53)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
               _StyleDefs(54)  =   "Named:id=33:Normal"
               _StyleDefs(55)  =   ":id=33,.parent=0"
               _StyleDefs(56)  =   "Named:id=34:Heading"
               _StyleDefs(57)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(58)  =   ":id=34,.wraptext=-1"
               _StyleDefs(59)  =   "Named:id=35:Footing"
               _StyleDefs(60)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(61)  =   "Named:id=36:Selected"
               _StyleDefs(62)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(63)  =   "Named:id=37:Caption"
               _StyleDefs(64)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(65)  =   "Named:id=38:HighlightRow"
               _StyleDefs(66)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(67)  =   "Named:id=39:EvenRow"
               _StyleDefs(68)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(69)  =   "Named:id=40:OddRow"
               _StyleDefs(70)  =   ":id=40,.parent=33"
               _StyleDefs(71)  =   "Named:id=41:RecordSelector"
               _StyleDefs(72)  =   ":id=41,.parent=34"
               _StyleDefs(73)  =   "Named:id=42:FilterBar"
               _StyleDefs(74)  =   ":id=42,.parent=33"
            End
            Begin VB.Shape SFocusTableTraçabiliteAlarmes 
               BorderColor     =   &H000000FF&
               BorderWidth     =   4
               Height          =   10710
               Left            =   285
               Top             =   405
               Visible         =   0   'False
               Width           =   27750
            End
         End
         Begin VB.Frame FCriteresRecherche 
            Caption         =   " Critères de recherche "
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
            Left            =   180
            TabIndex        =   9
            Top             =   120
            Width           =   28335
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
               Height          =   375
               Left            =   13740
               TabIndex        =   2
               Top             =   300
               Width           =   4215
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
               ItemData        =   "FTraçabiliteAlarmes.frx":2AE0E
               Left            =   1680
               List            =   "FTraçabiliteAlarmes.frx":2AE1B
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   300
               Width           =   4215
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
               Height          =   375
               Left            =   7620
               TabIndex        =   0
               Top             =   300
               Width           =   4215
            End
            Begin VB.CommandButton CBRaz 
               BackColor       =   &H00C0C0FF&
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
               Height          =   375
               Left            =   12000
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   " Annule tris et recherches "
               Top             =   300
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Contenant"
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
               Left            =   12720
               TabIndex        =   15
               Top             =   360
               Width           =   900
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Commençant par"
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
               Index           =   30
               Left            =   6015
               TabIndex        =   11
               Top             =   360
               Width           =   1545
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Rechercher par"
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
               Index           =   29
               Left            =   180
               TabIndex        =   10
               Top             =   360
               Width           =   1395
            End
         End
      End
   End
End
Attribute VB_Name = "FTraçabiliteAlarmes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant la traçabilité des alarmes
' Nom                    : FTraçabiliteAlarmes.frm
' Date de création : 27/10/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "TRACABILITE DES ALARMES"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum IDX_RECHERCHER_PAR
    IDX_DATE_DETECTION_DEFAUT = 1
    IDX_NUM_DEFAUT = 2
    IDX_DATE_CORRECTION_DEFAUT = 3
End Enum

Private Enum COLONNES_DETAILS_TRACABILITE_ALARMES
    C_NUM_DEFAUT = 0
    C_DATE_DETECTION_DEFAUT = 1
    C_DATE_CORRECTION_DEFAUT = 2
    C_LIBELLE_DEFAUT = 3
    C_COMPLEMENT_DEFAUT = 4
End Enum

'--- types privées ---

'--- Variables privées
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean        'pour interdire certains évènements
Private LigneDepartDeplacement As Integer   'ligne de départ en cas de déplacement d'un détail
Private LigneArriveeDeplacement As Integer  'ligne de d'arrivée en cas de déplacement d'un détail
Private MemDernierBouton As Long                'mémoire du dernier bouton

'--- tableaux privés ---

'--- Variables publiques
Public NumFenetre As Long                              'numéro de la fenêtre lorsqu'elle devient active

Private Sub ADODCtraçabilitealarmes_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    With pRecordset
        
        If .BOF = False And .EOF = False Then
        
            '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
            Select Case MemDernierBouton
                Case ETATS_BOUTONS.E_AVANT_NOUVEAU, ETATS_BOUTONS.E_APRES_NOUVEAU
                    Me.Caption = UCase(TITRE_FENETRE) & " - "
                    LRenseignements.Caption = "-"
                Case Else
                    Me.Caption = UCase(TITRE_FENETRE) & " - Défaut n° " & pRecordset("NumDefaut") & _
                                                                        " (" & pRecordset("LibelleDefaut") & ") du " & _
                                                                        pRecordset("DateDetectionDefaut")
                    LRenseignements.Caption = .AbsolutePosition & "/" & .RecordCount
            End Select
       
        Else
       
            '--- si fiche inexistante affichage d'un tiret ---
            Me.Caption = UCase(TITRE_FENETRE)
            LRenseignements.Caption = "-"
       
        End If
    
        '--- affichage des renseignements de la Fenetre ---
        LRenseignementsFenetre.Caption = Me.Caption
    
    End With
    
End Sub

Private Sub CBActualiser_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_ACTUALISER
    
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- actualisation ---
    ADODCTraçabiliteAlarmes.Refresh
    
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

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBQuitter_Click()
    On Error Resume Next
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

Private Sub CBRaz_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- vidage des champs / lancement de la requête ---
    With TBCommencantPar
        .Text = ""
        .Refresh
        .SetFocus
    End With
    With TBContenant
        .Text = ""
        .Refresh
    End With
    LanceRecherche

End Sub

Private Sub CBRechercherPar_Click()
    On Error Resume Next
    If PremiereActivation = True Then
        DoEvents
        CBRaz_Click
    End If
End Sub

Private Sub CBSupprimer_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- demande de confirmation ---
    If AppelFenetre(F_MESSAGE, _
                            TITRE_MESSAGES, _
                            vbCrLf & vbCrLf & _
                            "La totalité de la traçabilité des alarmes sera effacée." & vbCrLf & _
                            vbCrLf & vbCrLf & _
                            "cs|Voulez-vous réellement tout effacer ?", _
                            2, 0, 1) = vbYes Then

        '--- appel de la routine ---
        SuppressionTotaliteTraçabiliteAlarmes
        CBActualiser_Click
    
    End If

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

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        PremiereActivation = True
        If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
        Me.Refresh
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des touches communes ---
    Call OccFSynoptique.GestionTouches(KeyCode, Shift)
    
    '--- gestion des touches relatifs à cette Fenetre ---
    GestionTouches KeyCode, Shift

End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
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

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    ADODCTraçabiliteAlarmes.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - ADODCTraçabiliteAlarmes.Width
    LRenseignements.Left = ADODCTraçabiliteAlarmes.Left
    CBActualiser.Left = ADODCTraçabiliteAlarmes.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
    CBSupprimer.Left = CBActualiser.Left - MARGES.M_ENTRE_BOUTONS - CBSupprimer.Width
    
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
            
                '--- calculs des emplacements ---
            
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
' Rôle      : Initialise la fenêtre ( ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- affectation ---
    PremiereActivation = False
    
    '--- divers sur la fenêtre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .WindowState = vbMaximized
    End With
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- renseignements de la fenêtre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)
    
    '--- gestion des détails ---
    GestionTraçabiliteAlarmes GG_INITIALISATION
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- Aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False

    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- dé de la Fenetre ---
    Me.Visible = False
    Unload Me
    Set OccFTraçabiliteAlarmes = Nothing
    
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

Private Sub TBContenant_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRecherche
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBContenant_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES
End Sub

Private Sub TDBGTraçabiliteAlarmes_GotFocus()
    On Error Resume Next
    SFocusTableTraçabiliteAlarmes.Visible = True
End Sub

Private Sub TDBGTraçabiliteAlarmes_LostFocus()
    On Error Resume Next
    SFocusTableTraçabiliteAlarmes.Visible = False
End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then          'obligation de passer par le bouton quitter
        Cancel = True
        CBQuitter_Click
    End If
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
            CBQuitter.Enabled = True
            ADODCTraçabiliteAlarmes.Enabled = True
            CBActualiser.Enabled = True
            CBQuitter.Enabled = True
            FCriteresRecherche.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
            ADODCTraçabiliteAlarmes.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            FCriteresRecherche.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
            ADODCTraçabiliteAlarmes.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            FCriteresRecherche.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
            ADODCTraçabiliteAlarmes.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            FCriteresRecherche.Enabled = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
            CBQuitter.Enabled = True
            ADODCTraçabiliteAlarmes.Enabled = False
            CBActualiser.Enabled = False
            FCriteresRecherche.Enabled = False

        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- après nouveau ---
            FCriteresRecherche.Enabled = False
            CBQuitter.Enabled = True
            ADODCTraçabiliteAlarmes.Enabled = False
            CBActualiser.Enabled = False
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
            ADODCTraçabiliteAlarmes.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            FCriteresRecherche.Enabled = True
        
        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

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

Private Sub TBCommencantPar_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRecherche
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBCommencantPar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case Succ(CBRechercherPar.ListIndex)
        Case IDX_RECHERCHER_PAR.IDX_DATE_DETECTION_DEFAUT: FiltreToucheASCII KeyAscii, DONNEES.D_DATE_JJMMAAAA        'date de détection du défaut
        Case IDX_RECHERCHER_PAR.IDX_NUM_DEFAUT: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 6                           'n° du défaut
        Case Else
    End Select
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lance une recherche en fonction des critères
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LanceRecherche()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
        
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
    CommencantPar = TBCommencantPar.Text
    Contenant = TBContenant.Text
    IdxRecherchePar = Succ(CBRechercherPar.ListIndex)
    If IdxRecherchePar < 1 Then IdxRecherchePar = 1
    RechercherPar = Choose(IdxRecherchePar, _
                                              "TraçabiliteAlarmes.DateDetectionDefaut", _
                                              "TraçabiliteAlarmes.NumDefaut", _
                                              "TraçabiliteAlarmes.DateCorrectionDefaut")

    '--- début de la requête ---
    RequeteSQL = "SELECT TraçabiliteAlarmes.*, ListeDefauts.LibelleDefaut AS LibelleDefaut " & _
                            "FROM TraçabiliteAlarmes LEFT OUTER JOIN ListeDefauts ON TraçabiliteAlarmes.NumDefaut = ListeDefauts.NumDefaut "
    
    If IdxRecherchePar = IDX_RECHERCHER_PAR.IDX_DATE_DETECTION_DEFAUT Or IdxRecherchePar = IDX_RECHERCHER_PAR.IDX_DATE_CORRECTION_DEFAUT Then
        
        '--- filtres pour la date ---
        Filtre1 = "(CONVERT(VARCHAR(10), " & RechercherPar & ", 103) LIKE '" & CommencantPar & "%')"
        Filtre2 = "(CONVERT(VARCHAR(10), " & RechercherPar & ", 103) LIKE '%" & Contenant & "%')"
    
    Else
        
        '--- filtres pour chaines de caractères ---
        Filtre1 = "(" & RechercherPar & " LIKE '" & CommencantPar & "%')"
        Filtre2 = "(" & RechercherPar & " LIKE '%" & Contenant & "%')"
    
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
    
    '--- fin de la requête ---
    RequeteSQL = RequeteSQL & "ORDER BY " & RechercherPar
    Select Case IdxRecherchePar
        Case IDX_RECHERCHER_PAR.IDX_DATE_DETECTION_DEFAUT, IDX_RECHERCHER_PAR.IDX_DATE_CORRECTION_DEFAUT
            RequeteSQL = RequeteSQL & " DESC, TraçabiliteAlarmes.NumDefaut"
        Case IDX_RECHERCHER_PAR.IDX_NUM_DEFAUT
            RequeteSQL = RequeteSQL & ", TraçabiliteAlarmes.DateDetectionDefaut DESC"
        Case Else
    End Select

    With ADODCTraçabiliteAlarmes
        
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
    'Debug.Print RequeteSQL
    
    '--- curseur de la souris ---
    SourisEnAttente False
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
    
    '--- curseur de la souris ---
    SourisEnAttente False
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées :                     RechercherPar -> Valeur du champ TBRechercherPar
'                                  CommencantPar -> Valeur du champ TBCommencantPar
'                                            Contenant -> Valeur du champ TBContenant
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal RechercherPar As Integer, _
                                                   ByVal CommencantPar As String, _
                                                   ByVal Contenant As String)


    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rechercher par ---
    If RechercherPar > 0 Then
        CBRechercherPar.ListIndex = Pred(RechercherPar)
    Else
        CBRechercherPar.Text = CBRechercherPar.List(0)
    End If
    
    '--- commençant par ---
    TBCommencantPar.Text = CommencantPar
    
    '--- contenant ---
    TBContenant.Text = Contenant
    
    '--- lancement de la recherche ---
    LanceRecherche

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Suppression de la totalité de la table de traçabilite des alarmes
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SuppressionTotaliteTraçabiliteAlarmes()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim Requete As String
    Dim ConnexionBDAnodisationSQL As New ADODB.Connection
    Dim Enregistrement As New ADODB.Recordset
    
    '--- ouverture de la connexion à la base de données d'anodisation en SQL SERVER 2000 ---
    With ConnexionBDAnodisationSQL
        .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        .CursorLocation = adUseServer
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    
    '--- lancement de la requête ---
    With Enregistrement
        .CursorLocation = adUseServer
        Requete = "DELETE FROM " & TABLE_TRACABILITE_ALARMES
        .Open Requete, ConnexionBDAnodisationSQL, adLockOptimistic, adCmdText
    End With
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion de la traçabilité des alarmes
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionTraçabiliteAlarmes(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---

    '--- affectation ---
    GestionTraçabiliteAlarmes = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGTraçabiliteAlarmes
                
                .Visible = False                                                           'rendre la grille invisible
                '.ClearFields                                                                'effacer la structure
            
                .Splits(0).AllowSizing = False                                      'division de la grille
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_3                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.JAUNE_3                     'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1      'couleur des lignes impaires
                
                .ForeColor = COULEURS.BLEU_4                                'couleurs des données
                
                .HeadFont.Name = "Arial"
                With .Font
                    .Name = "MS Sans serif"
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.05
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
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
                
                With .Columns(COLONNES_DETAILS_TRACABILITE_ALARMES.C_NUM_DEFAUT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du défaut"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_DETAILS_TRACABILITE_ALARMES.C_DATE_DETECTION_DEFAUT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Date de détection du défaut"
                    .Width = EPAISSEUR_CARACTERE * 25
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_TRACABILITE_ALARMES.C_DATE_CORRECTION_DEFAUT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Date de correction du défaut"
                    .Width = EPAISSEUR_CARACTERE * 25
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_DETAILS_TRACABILITE_ALARMES.C_LIBELLE_DEFAUT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé du défaut"
                    .Width = EPAISSEUR_CARACTERE * 72.5
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_DETAILS_TRACABILITE_ALARMES.C_COMPLEMENT_DEFAUT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Informations complémentaires sur le défaut"
                    .Width = EPAISSEUR_CARACTERE * 60
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
    GestionTraçabiliteAlarmes = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gère l'appui des touches du clavier
' Entrées :
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


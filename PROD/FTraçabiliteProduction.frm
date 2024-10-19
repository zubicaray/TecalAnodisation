VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FTraçabiliteProduction 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   13020
   ClientLeft      =   -405
   ClientTop       =   4890
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "Marlett"
      Size            =   7.5
      Charset         =   2
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FTraçabiliteProduction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13020
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FTraçabiliteProduction.frx":0442
      ScaleHeight     =   315
      ScaleWidth      =   13620
      TabIndex        =   5
      Top             =   0
      Width           =   13680
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "CHARGE GEREE"
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
         TabIndex        =   6
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
      ScaleWidth      =   13620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   11925
      Width           =   13680
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   40
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   39
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FTraçabiliteProduction.frx":24D84
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
            Picture         =   "FTraçabiliteProduction.frx":24F2E
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FTraçabiliteProduction.frx":250D8
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
         Left            =   18180
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTraçabiliteProduction.frx":257DA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualise&r"
         DownPicture     =   "FTraçabiliteProduction.frx":25EDC
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
         Left            =   13680
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTraçabiliteProduction.frx":265DE
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   " Actualiser les données "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin MSAdodcLib.Adodc ADODCDetailsChargesProduction 
         Height          =   435
         Left            =   15420
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
         MaxRecords      =   500
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
         RecordSource    =   $"FTraçabiliteProduction.frx":26CE0
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
         Left            =   2040
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
               Picture         =   "FTraçabiliteProduction.frx":26E5C
               Key             =   "fleche noire"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27068
               Key             =   "fleche blanche"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27274
               Key             =   "fleche grise"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27480
               Key             =   "fleche rouge"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":2768C
               Key             =   "fleche jaune"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27898
               Key             =   "fleche verte"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27AA4
               Key             =   "fleche cyan"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27CB0
               Key             =   "fleche bleue"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":27EBC
               Key             =   "etoile noire"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":280C8
               Key             =   "etoile blanche"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":282D4
               Key             =   "etoile grise"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":284E0
               Key             =   "etoile rouge"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":286EC
               Key             =   "etoile jaune"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":288F8
               Key             =   "etoile verte"
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":28B04
               Key             =   "etoile cyan"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":28D10
               Key             =   "etoile bleue"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":28F1C
               Key             =   "modification noire"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29120
               Key             =   "modification blanche"
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29324
               Key             =   "modification grise"
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29528
               Key             =   "modification rouge"
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":2972C
               Key             =   "modification jaune"
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29930
               Key             =   "modification vert"
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29B34
               Key             =   "modification cyan"
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29D38
               Key             =   "modification bleue"
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":29F3C
               Key             =   "indicateur vert"
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FTraçabiliteProduction.frx":2A140
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   315
         Left            =   1440
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
         Left            =   15420
         TabIndex        =   3
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12915
      Index           =   0
      Left            =   0
      ScaleHeight     =   12915
      ScaleWidth      =   13680
      TabIndex        =   1
      Top             =   375
      Width           =   13680
      Begin VB.PictureBox PBDeplacementFenetre 
         Height          =   12705
         Index           =   1
         Left            =   0
         ScaleHeight     =   12645
         ScaleWidth      =   28635
         TabIndex        =   2
         Top             =   0
         Width           =   28695
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
            Height          =   1995
            Left            =   240
            ScaleHeight     =   1965
            ScaleWidth      =   28185
            TabIndex        =   41
            Top             =   240
            Width           =   28215
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
               TabIndex        =   57
               Top             =   600
               Width           =   2655
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
               TabIndex        =   56
               Top             =   180
               Width           =   2655
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
               ItemData        =   "FTraçabiliteProduction.frx":2A344
               Left            =   1740
               List            =   "FTraçabiliteProduction.frx":2A354
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   480
               Width           =   3375
            End
            Begin VB.TextBox TBCritereEntre2Valeurs 
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
               Left            =   13680
               TabIndex        =   49
               Top             =   600
               Width           =   2655
            End
            Begin VB.TextBox TBCritereEntre2Valeurs 
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
               Left            =   13680
               TabIndex        =   48
               Top             =   180
               Width           =   2655
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
               Picture         =   "FTraçabiliteProduction.frx":2A3A7
               Style           =   1  'Graphical
               TabIndex        =   47
               ToolTipText     =   " Lancer une recherche "
               Top             =   180
               UseMaskColor    =   -1  'True
               Width           =   1335
            End
            Begin VB.CommandButton CBRechercherSurGrille 
               BackColor       =   &H00E0E0E0&
               Caption         =   "GRILLE"
               DownPicture     =   "FTraçabiliteProduction.frx":2A6E9
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
               Picture         =   "FTraçabiliteProduction.frx":2ADEB
               Style           =   1  'Graphical
               TabIndex        =   46
               TabStop         =   0   'False
               ToolTipText     =   " Rechercher sur la grille "
               Top             =   180
               UseMaskColor    =   -1  'True
               Width           =   915
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
               Picture         =   "FTraçabiliteProduction.frx":2B4ED
               Style           =   1  'Graphical
               TabIndex        =   45
               ToolTipText     =   " Annule tris et recherches "
               Top             =   180
               UseMaskColor    =   -1  'True
               Width           =   915
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
               Top             =   180
               Value           =   -1  'True
               Width           =   375
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
               TabIndex        =   42
               ToolTipText     =   " Change la présentation de la grille "
               Top             =   540
               Width           =   375
            End
            Begin TrueOleDBGrid80.TDBGrid TDBGGrilleRecherche 
               Bindings        =   "FTraçabiliteProduction.frx":2B6DF
               Height          =   10875
               Left            =   240
               TabIndex        =   44
               Top             =   1140
               Width           =   27675
               _ExtentX        =   48816
               _ExtentY        =   19182
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "NumCommandeInterne"
               Columns(0).DataField=   "NumCommandeInterne"
               Columns(0).DataWidth=   8
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "NbrReparations"
               Columns(1).DataField=   "NbrReparations"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "NumFicheProduction"
               Columns(2).DataField=   "NumFicheProduction"
               Columns(2).DataWidth=   8
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "DateEntreeEnLigne"
               Columns(3).DataField=   "DateEntreeEnLigne"
               Columns(3).DataWidth=   19
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "DateArriveeAuDechargement"
               Columns(4).DataField=   "DateArriveeAuDechargement"
               Columns(4).DataWidth=   19
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "CodeClient"
               Columns(5).DataField=   "CodeClient"
               Columns(5).DataWidth=   8
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "NbrPieces"
               Columns(6).DataField=   "NbrPieces"
               Columns(6).DataWidth=   11
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "Designation"
               Columns(7).DataField=   "Designation"
               Columns(7).DataWidth=   255
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).Caption=   "Matiere"
               Columns(8).DataField=   "Matiere"
               Columns(8).DataWidth=   30
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(9)._VlistStyle=   0
               Columns(9)._MaxComboItems=   5
               Columns(9).Caption=   "NumGammeAnodisation"
               Columns(9).DataField=   "NumGammeAnodisation"
               Columns(9).DataWidth=   6
               Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(10)._VlistStyle=   0
               Columns(10)._MaxComboItems=   5
               Columns(10).Caption=   "RefGammeAnodisation"
               Columns(10).DataField=   "RefGammeAnodisation"
               Columns(10).DataWidth=   18
               Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(11)._VlistStyle=   0
               Columns(11)._MaxComboItems=   5
               Columns(11).Caption=   "ChargePrioritaireTexte"
               Columns(11).DataField=   "ChargePrioritaireTexte"
               Columns(11).DataWidth=   3
               Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   12
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   503
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).DividerColor=   13160660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=12"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=4471"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4339"
               Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
               Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
               Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(9)=   "Column(2).Width=4366"
               Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
               Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(13)=   "Column(3).Width=4154"
               Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4022"
               Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(17)=   "Column(3)._AlignLeft=0"
               Splits(0)._ColumnProps(18)=   "Column(4).Width=5689"
               Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=5556"
               Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(22)=   "Column(4)._AlignLeft=0"
               Splits(0)._ColumnProps(23)=   "Column(5).Width=2355"
               Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2223"
               Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(27)=   "Column(6).Width=2461"
               Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2328"
               Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(31)=   "Column(6)._AlignLeft=0"
               Splits(0)._ColumnProps(32)=   "Column(7).Width=4366"
               Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=4233"
               Splits(0)._ColumnProps(35)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(36)=   "Column(8).Width=4366"
               Splits(0)._ColumnProps(37)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(38)=   "Column(8)._WidthInPix=4233"
               Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
               Splits(0)._ColumnProps(40)=   "Column(9).Width=4736"
               Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
               Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=4604"
               Splits(0)._ColumnProps(43)=   "Column(9).Order=10"
               Splits(0)._ColumnProps(44)=   "Column(10).Width=4551"
               Splits(0)._ColumnProps(45)=   "Column(10).DividerColor=0"
               Splits(0)._ColumnProps(46)=   "Column(10)._WidthInPix=4419"
               Splits(0)._ColumnProps(47)=   "Column(10).Order=11"
               Splits(0)._ColumnProps(48)=   "Column(11).Width=4366"
               Splits(0)._ColumnProps(49)=   "Column(11).DividerColor=0"
               Splits(0)._ColumnProps(50)=   "Column(11)._WidthInPix=4233"
               Splits(0)._ColumnProps(51)=   "Column(11).Order=12"
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
               _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
               _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
               _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
               _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
               _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
               _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
               _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
               _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
               _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
               _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
               _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
               _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
               _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=78,.parent=13"
               _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=75,.parent=14"
               _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=76,.parent=15"
               _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=77,.parent=17"
               _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=82,.parent=13"
               _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=14"
               _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=15"
               _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=17"
               _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=28,.parent=13"
               _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
               _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
               _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
               _StyleDefs(84)  =   "Named:id=33:Normal"
               _StyleDefs(85)  =   ":id=33,.parent=0"
               _StyleDefs(86)  =   "Named:id=34:Heading"
               _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(88)  =   ":id=34,.wraptext=-1"
               _StyleDefs(89)  =   "Named:id=35:Footing"
               _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(91)  =   "Named:id=36:Selected"
               _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(93)  =   "Named:id=37:Caption"
               _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(95)  =   "Named:id=38:HighlightRow"
               _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(97)  =   "Named:id=39:EvenRow"
               _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(99)  =   "Named:id=40:OddRow"
               _StyleDefs(100) =   ":id=40,.parent=33"
               _StyleDefs(101) =   "Named:id=41:RecordSelector"
               _StyleDefs(102) =   ":id=41,.parent=34"
               _StyleDefs(103) =   "Named:id=42:FilterBar"
               _StyleDefs(104) =   ":id=42,.parent=33"
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Entre"
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
               Index           =   10
               Left            =   12960
               TabIndex        =   54
               Top             =   240
               Width           =   555
            End
            Begin VB.Label LLibelles 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "et"
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
               Index           =   9
               Left            =   13320
               TabIndex        =   53
               Top             =   660
               Width           =   210
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
               Index           =   8
               Left            =   6540
               TabIndex        =   52
               Top             =   660
               Width           =   1050
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
               Height          =   240
               Index           =   7
               Left            =   6540
               TabIndex        =   51
               Top             =   240
               Width           =   1755
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
               Index           =   6
               Left            =   1740
               TabIndex        =   50
               Top             =   180
               Width           =   3360
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
         End
         Begin C1SizerLibCtl.C1Tab CTOnglets 
            Height          =   10935
            Left            =   240
            TabIndex        =   7
            Top             =   1560
            Width           =   28215
            _cx             =   49768
            _cy             =   19288
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
            Caption         =   "Détails de la charge|Gamme d'ANODISATION|Traçabilité de la charge|Alarmes de la ligne"
            Align           =   0
            CurrTab         =   1
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
            Picture(0)      =   "FTraçabiliteProduction.frx":2B70B
            Picture(1)      =   "FTraçabiliteProduction.frx":2B865
            Picture(2)      =   "FTraçabiliteProduction.frx":2B9BF
            Picture(3)      =   "FTraçabiliteProduction.frx":2BB19
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   3
               Left            =   29160
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   11
               Top             =   495
               Width           =   28125
               Begin VB.TextBox TBAlarmesLigne 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   9855
                  Left            =   240
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   35
                  Top             =   240
                  Width           =   27615
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   2
               Left            =   28860
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   17
               Top             =   495
               Width           =   28125
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsFichesProduction 
                  Height          =   9840
                  Left            =   240
                  TabIndex        =   34
                  Top             =   240
                  Width           =   27615
                  _ExtentX        =   48710
                  _ExtentY        =   17357
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   0
                  Cols            =   7
                  BackColorFixed  =   33023
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
                  GridLinesUnpopulated=   3
                  AllowUserResizing=   3
                  Appearance      =   0
                  GridLineWidthUnpopulated=   2
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
                  _NumberOfBands  =   2
                  _Band(0).Cols   =   7
                  _Band(0).GridLinesBand=   1
                  _Band(0).TextStyleBand=   0
                  _Band(0).TextStyleHeader=   0
                  _Band(1).BandIndent=   1
                  _Band(1).Cols   =   4
                  _Band(1).GridLinesBand=   1
                  _Band(1).TextStyleBand=   0
                  _Band(1).TextStyleHeader=   0
                  _Band(1).ColHeader=   1
               End
               Begin VB.Shape SFocusTableDetailsFichesProduction 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   9855
                  Left            =   240
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   27630
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   8
               Left            =   30960
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   16
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   7
               Left            =   30660
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   15
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   6
               Left            =   30360
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   14
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   5
               Left            =   30060
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   13
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   4
               Left            =   29460
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   12
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   0
               Left            =   -28770
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   10
               Top             =   495
               Width           =   28125
               Begin VB.Frame FControles 
                  Caption         =   " Contrôles "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   9975
                  Left            =   18540
                  TabIndex        =   58
                  Top             =   120
                  Width           =   9375
                  Begin VB.Label LControleObservations 
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
                     Height          =   555
                     Left            =   2520
                     TabIndex        =   68
                     Top             =   1740
                     Width           =   6615
                  End
                  Begin VB.Label LControleColoration 
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
                     Left            =   2520
                     TabIndex        =   67
                     Top             =   1260
                     Width           =   6615
                  End
                  Begin VB.Label LControleEpaisseurAnodisation 
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
                     Left            =   2520
                     TabIndex        =   66
                     Top             =   840
                     Width           =   735
                  End
                  Begin VB.Label LControleColmatage 
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
                     Left            =   2520
                     TabIndex        =   65
                     Top             =   420
                     Width           =   735
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Observations"
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
                     Index           =   14
                     Left            =   180
                     TabIndex        =   62
                     Top             =   1740
                     Width           =   2175
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Coloration"
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
                     Index           =   13
                     Left            =   480
                     TabIndex        =   61
                     Top             =   1320
                     Width           =   1875
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Epaisseur d'anodisation"
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
                     Index           =   12
                     Left            =   180
                     TabIndex        =   60
                     Top             =   900
                     Width           =   2175
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Colmatage"
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
                     Index           =   11
                     Left            =   180
                     TabIndex        =   59
                     Top             =   480
                     Width           =   2175
                     WordWrap        =   -1  'True
                  End
               End
               Begin VB.Frame FRenseignements 
                  Caption         =   " Renseignements de production pour cette charge "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1635
                  Left            =   180
                  TabIndex        =   18
                  Top             =   120
                  Width           =   18135
                  Begin VB.Label LNumBarre 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H80000005&
                     BorderStyle     =   1  'Fixed Single
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   15.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   375
                     Left            =   9480
                     TabIndex        =   72
                     Top             =   720
                     Width           =   615
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "N° de la barre"
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
                     Index           =   16
                     Left            =   7920
                     TabIndex        =   71
                     Top             =   780
                     Width           =   1395
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Date d'entrée en ligne"
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
                     Index           =   2
                     Left            =   420
                     TabIndex        =   26
                     Top             =   780
                     Width           =   2655
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LDateEntreeEnLigne 
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
                     Left            =   3240
                     TabIndex        =   25
                     Top             =   720
                     Width           =   2655
                  End
                  Begin VB.Label LNumFicheProduction 
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
                     Left            =   3240
                     TabIndex        =   24
                     Top             =   300
                     Width           =   1455
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "N° de la fiche de production"
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
                     Left            =   480
                     TabIndex        =   23
                     Top             =   360
                     Width           =   2595
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LChargePrioritaire 
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
                     Left            =   9480
                     TabIndex        =   22
                     Top             =   300
                     Width           =   735
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Charge prioritaire dans la production"
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
                     Index           =   5
                     Left            =   5760
                     TabIndex        =   21
                     Top             =   360
                     Width           =   3555
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LDateArriveeAuDechargement 
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
                     Left            =   3240
                     TabIndex        =   20
                     Top             =   1140
                     Width           =   2655
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Date d'arrivée au déchargement"
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
                     Index           =   1
                     Left            =   120
                     TabIndex        =   19
                     Top             =   1200
                     Width           =   2955
                     WordWrap        =   -1  'True
                  End
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsChargesProduction 
                  Height          =   8115
                  Left            =   240
                  TabIndex        =   27
                  Top             =   1980
                  Width           =   18075
                  _ExtentX        =   31882
                  _ExtentY        =   14314
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   0
                  Rows            =   100
                  Cols            =   6
                  BackColorFixed  =   8421376
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
               Begin VB.Shape SFocusTableDetailsChargesProduction 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   8130
                  Left            =   240
                  Top             =   1980
                  Visible         =   0   'False
                  Width           =   18090
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   1
               Left            =   45
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   9
               Top             =   495
               Width           =   28125
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
                  Height          =   4335
                  Left            =   15540
                  TabIndex        =   73
                  Top             =   2400
                  Width           =   12375
                  Begin VB.CommandButton CBVisualisationGraphesProduction 
                     BackColor       =   &H00C0FFC0&
                     Caption         =   "Visualisation du graphe"
                     DownPicture     =   "FTraçabiliteProduction.frx":2BC73
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
                     Left            =   240
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FTraçabiliteProduction.frx":2C375
                     Style           =   1  'Graphical
                     TabIndex        =   108
                     Top             =   3060
                     UseMaskColor    =   -1  'True
                     Width           =   2955
                  End
                  Begin VB.PictureBox PBPhasesRedresseurs 
                     BackColor       =   &H00C0E0FF&
                     Height          =   3735
                     Left            =   4560
                     ScaleHeight     =   3675
                     ScaleWidth      =   6015
                     TabIndex        =   74
                     Top             =   360
                     Width           =   6075
                     Begin VB.Label LIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   4
                        Left            =   4440
                        TabIndex        =   107
                        Top             =   2460
                        Width           =   855
                     End
                     Begin VB.Label LIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   3
                        Left            =   4440
                        TabIndex        =   106
                        Top             =   1920
                        Width           =   855
                     End
                     Begin VB.Label LIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   2
                        Left            =   4440
                        TabIndex        =   105
                        Top             =   1380
                        Width           =   855
                     End
                     Begin VB.Label LIntensitesPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   1
                        Left            =   4440
                        TabIndex        =   104
                        Top             =   840
                        Width           =   855
                     End
                     Begin VB.Label LTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   4
                        Left            =   2880
                        TabIndex        =   103
                        Top             =   2460
                        Width           =   855
                     End
                     Begin VB.Label LTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   3
                        Left            =   2880
                        TabIndex        =   102
                        Top             =   1920
                        Width           =   855
                     End
                     Begin VB.Label LTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   2
                        Left            =   2880
                        TabIndex        =   101
                        Top             =   1380
                        Width           =   855
                     End
                     Begin VB.Label LTensionsPhases 
                        Alignment       =   1  'Right Justify
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
                        Index           =   1
                        Left            =   2880
                        TabIndex        =   100
                        Top             =   840
                        Width           =   855
                     End
                     Begin VB.Label LTempsPhases 
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
                        Index           =   4
                        Left            =   1560
                        TabIndex        =   99
                        Top             =   2460
                        Width           =   855
                     End
                     Begin VB.Label LTempsPhases 
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
                        Index           =   3
                        Left            =   1560
                        TabIndex        =   98
                        Top             =   1920
                        Width           =   855
                     End
                     Begin VB.Label LTempsPhases 
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
                        Index           =   2
                        Left            =   1560
                        TabIndex        =   97
                        Top             =   1380
                        Width           =   855
                     End
                     Begin VB.Label LTempsPhases 
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
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   96
                        Top             =   840
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
                        Index           =   10
                        X1              =   4200
                        X2              =   4200
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
                        Index           =   12
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
                        Index           =   18
                        Left            =   1440
                        TabIndex        =   89
                        Top             =   360
                        Width           =   1095
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
                        Index           =   17
                        Left            =   480
                        TabIndex        =   88
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
                        TabIndex        =   87
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
                        TabIndex        =   86
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
                        TabIndex        =   85
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
                        Index           =   24
                        Left            =   5400
                        TabIndex        =   84
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
                        Index           =   25
                        Left            =   5400
                        TabIndex        =   83
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
                        Index           =   26
                        Left            =   5400
                        TabIndex        =   82
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
                        Index           =   27
                        Left            =   4320
                        TabIndex        =   81
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
                        Index           =   28
                        Left            =   2760
                        TabIndex        =   80
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
                        Index           =   29
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
                        Index           =   30
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
                        Index           =   31
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
                        Index           =   32
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
                        Index           =   35
                        Left            =   3840
                        TabIndex        =   75
                        Top             =   870
                        Width           =   195
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
                        Index           =   3
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   720
                        Width           =   5535
                     End
                     Begin VB.Shape SDecoration 
                        FillColor       =   &H00FFFFC0&
                        FillStyle       =   0  'Solid
                        Height          =   555
                        Index           =   5
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1260
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
                        Index           =   6
                        Left            =   240
                        Shape           =   4  'Rounded Rectangle
                        Top             =   1800
                        Width           =   5535
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
                     TabIndex        =   95
                     Top             =   1200
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
                     TabIndex        =   94
                     Top             =   780
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
                     TabIndex        =   93
                     Top             =   360
                     Width           =   2910
                  End
                  Begin VB.Image IPhasesAnodisation 
                     Height          =   2010
                     Left            =   240
                     Picture         =   "FTraçabiliteProduction.frx":2CA77
                     Top             =   660
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
                     TabIndex        =   92
                     Top             =   360
                     Width           =   1170
                  End
                  Begin VB.Shape SDecoration 
                     BorderWidth     =   2
                     FillColor       =   &H00FFFFC0&
                     FillStyle       =   0  'Solid
                     Height          =   960
                     Index           =   8
                     Left            =   3150
                     Top             =   675
                     Width           =   1170
                  End
               End
               Begin VB.Frame FGammeAnodisation 
                  Caption         =   " Caractéristiques de la gamme "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   2115
                  Left            =   15540
                  TabIndex        =   28
                  Top             =   180
                  Width           =   12375
                  Begin VB.TextBox TBMatieresConcernees 
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   615
                     Left            =   2220
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   69
                     Top             =   1260
                     Width           =   7155
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Matières concernées"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   285
                     Index           =   4
                     Left            =   60
                     TabIndex        =   70
                     Top             =   1320
                     Width           =   2055
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Référence"
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
                     Index           =   15
                     Left            =   3540
                     TabIndex        =   64
                     Top             =   420
                     Width           =   975
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LRefGamme 
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
                     Left            =   4620
                     TabIndex        =   63
                     Top             =   360
                     Width           =   4755
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Gamme n°"
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
                     Index           =   0
                     Left            =   1140
                     TabIndex        =   31
                     Top             =   420
                     Width           =   975
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LNumGamme 
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
                     Left            =   2220
                     TabIndex        =   30
                     Top             =   360
                     Width           =   1155
                  End
                  Begin VB.Label LNomGamme 
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
                     Left            =   2220
                     TabIndex        =   29
                     Top             =   780
                     Width           =   7170
                  End
               End
               Begin MSMask.MaskEdBox MEBEditionDetailsGammesAnodisation 
                  Height          =   315
                  Left            =   540
                  TabIndex        =   32
                  Top             =   660
                  Visible         =   0   'False
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  BorderStyle     =   0
                  ClipMode        =   1
                  Appearance      =   0
                  BackColor       =   12632319
                  PromptInclude   =   0   'False
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
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsGammesProduction 
                  Height          =   9780
                  Left            =   240
                  TabIndex        =   33
                  Top             =   300
                  Width           =   15075
                  _ExtentX        =   26591
                  _ExtentY        =   17251
                  _Version        =   393216
                  BackColor       =   16777215
                  ForeColor       =   12582912
                  Rows            =   31
                  Cols            =   6
                  BackColorFixed  =   16576
                  ForeColorFixed  =   16777215
                  BackColorBkg    =   12648447
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
               Begin VB.Shape SFocusTableDetailsGammesProduction 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   9795
                  Left            =   240
                  Top             =   300
                  Visible         =   0   'False
                  Width           =   15090
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   10395
               Index           =   9
               Left            =   29760
               ScaleHeight     =   10335
               ScaleWidth      =   28065
               TabIndex        =   8
               Top             =   495
               Width           =   28125
            End
         End
      End
   End
End
Attribute VB_Name = "FTraçabiliteProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant la traçabilité de la production
' Nom                    : FTraçabiliteProduction.frm
' Date de création : 18/04/2002
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const NBR_COLONNES_DETAILS_CHARGES_PRODUCTION  As Integer = 7
Private Const NBR_COLONNES_DETAILS_GAMMES_PRODUCTION  As Integer = 6
Private Const NBR_COLONNES_DETAILS_FICHES_PRODUCTION  As Integer = 7

Private Const TITRE_FENETRE As String = "TRACABILITE DE LA PRODUCTION"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum ONGLETS
    O_DETAILS_CHARGE = 0
    O_GAMME_ANODISATION = 1
    O_DETAILS_FICHE_PRODUCTION = 2
    O_ALARMES_LIGNE = 3
End Enum

Private Enum COLONNES_GRILLE_RECHERCHE
    
    C_NUM_COMMANDE_INTERNE = 0
    
    C_NBR_REPARATIONS = 1                         'nombre de réparations
    
    C_NUM_FICHE_PRODUCTION = 2
    
    C_DATE_ENTREE_LIGNE = 3
    C_DATE_ARRIVEE_DECHARGEMENT = 4
    
    C_CODE_CLIENT = 5
    
    C_NBR_PIECES = 6
    C_DESIGNATION = 7
    C_MATIERE = 8
    
    C_NUM_GAMME = 9
    C_REF_GAMME = 10
    
    C_CHARGE_PRIORITAIRE = 11

End Enum

Private Enum IDX_RECHERCHER_PAR
    IDX_NUM_FICHE_PRODUCTION = 1
    IDX_NUM_COMMANDE_INTERNE = 2
    IDX_DATE_ENTREE_LIGNE = 3
    IDX_CODE_CLIENT = 4
End Enum

Private Enum COLONNES_DETAILS_CHARGES_PRODUCTION
    C_NUM_LIGNES = 0
    C_NUM_COMMANDE_INTERNE = 1
    C_NBR_REPARATIONS = 2                         'nombre de réparations
    C_CODE_CLIENT = 3
    C_NBR_PIECES = 4
    C_DESIGNATION = 5
    C_NUM_LIGNES_REFERENCES_CLIENT = 6
    C_MATIERE = 7
End Enum

Private Enum COLONNES_DETAILS_GAMMES_PRODUCTION
    C_NUM_LIGNES = 0
    C_CODE_ZONE = 1
    C_LIBELLE_ZONE = 2
    C_NOM_POSTE_REEL = 3
    C_TEMPS_AU_POSTE_TEXTE = 4
    C_DECOMPTE_TEMPS_POSTE_REEL = 5
    C_TEMPS_EGOUTTAGE_TEXTE = 6
End Enum

Private Enum COLONNES_DETAILS_FICHES_PRODUCTION
    C_NUM_LIGNES = 0
    C_NOM_POSTE = 1
    C_TEMPS_REEL_POSTE = 2
    C_TEMPS_REEL_EGOUTTAGE = 3
    C_TEMPERATURES = 4
    C_REDRESSEUR = 5
    C_ANALYSEUR = 6
    C_ALARMES_POSTE = 7
End Enum

'--- types privées ---
Private Type ImgDetailsChargesProduction
    NumCommandeInterne As Long                        'n° de commande interne
    NbrReparations As String                                    'nombre de réparations (champ texte volontaire)
    DateEntreeEnLigne As Date                                'date d'entrée en ligne
    DateArriveeAuDechargement As Date                 'date d'arrivée au déchargement
    NumBarre As Integer                                           'n° de barre
    CodeClient As String                                            'code client
    NbrPieces As String                                             'nombre de pièces
    Designation As String                                          'désignation
    NumLignesReferencesClient As String               'n° de lignes des références du client correspondant
                                                                                 'aux n° de lignes des travaux avec les quantités séparés par des tirets
    NbrLignesReferencesClient As Integer               'nombre de lignes des références du client une fois extraites

    NumGammeAnodisation As String                      'n° de la gamme d'anodisation
    RefGammeAnodisation As String                        'référence de la gamme d'anodisation
    Matiere As String                                                 'matière des pièces
    
    NumFicheProduction As String                            'n° de la fiche de production
    ChargePrioritaire As Boolean                              'indique qu'il s'agit d'une charge prioritaire
    AlarmesLigne As String                                       'alarmes de la ligne
    ControleColmatage As Integer                             'contrôle du colmatage (valeur de 0 à 5)
    ControleEpaisseurAnodisation As Integer           'contrôle de l'épaisseur d'anodisation (valeur de 0 à 100)
    ControleColoration As String                                'contrôle de la coloration (20 caractères)
    ControleObservations As String                           'observations sur les contrôles (50 caractères)
End Type

Private Type ImgDetailsGammesProduction
    NumZone As Integer                                             'n° de la zone
    Codezone As String                                              'code de la zone
    LibelleZone As String                                           'libellé de la zone
    TempsAuPosteTexte As String                             'temps au poste en texte au format HH:MM:SS
    TempsEgouttageTexte As String                          'temps d'égouttage en texte au format MM:SS
    TempsAuPosteSecondes As Long                        'temps au poste en secondes
    TempsEgouttageSecondes As Integer                  'temps d'égouttage en secondes
    NomPosteReel As String                                       'nom du poste réel (cas des postes multiples)
    DecompteDuTempsAuPosteReelTexte As String  'décompte du temps au poste réel en texte (HH:MM:SS)
End Type

Private Type ImgDetailsPhasesProduction
    ModeUouI As MODES_U_OU_I                            'mode tension ou intensité
    TempsPhase As Integer                                       'temps de la phase
    UPhase As Single                                                'tension de production
    IPhase As Single                                                  'intensité de production
End Type

Private Type ImgDetailsFichesProduction
    NumPoste As Integer                         'numéro du poste
    NomPoste As String                           'nom du poste
    TempsReelPoste As String                'temps réel au poste en HH:MM:SS
    TempsReelEgouttage As String         'temps d'égouttage en HH:MM:SS
    Temperatures As String                     'températures en entrée et sortie de bain
    Redresseur As String                        'U et I du redresseur
    Analyseur As String                           'analyseur en entrée et sortie du bain d'anodisation
    AlarmesPoste As String                     'Alarmes au poste
End Type

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean        'pour interdire certains évènements
Private LigneDepartDeplacement As Integer   'ligne de départ en cas de déplacement d'un détail
Private LigneArriveeDeplacement As Integer  'ligne de d'arrivée en cas de déplacement d'un détail
Private MemDernierBouton As Long                'mémoire du dernier bouton

'--- tableaux privés ---
Private TDetailsChargesProduction(1 To NBR_LIGNES_DETAILS_CHARGES) As ImgDetailsChargesProduction
Private TDetailsGammesProduction(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesProduction
Private TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As ImgDetailsPhasesProduction
Private TDetailsFichesProduction(1 To NBR_LIGNES_DETAILS_FICHES_PRODUCTION) As ImgDetailsFichesProduction

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active
Public RechercherSurGrille As Boolean          'publique pour le copier / coller

Private Sub ADODCDetailsChargesProduction_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    With pRecordset
        
        If .BOF = False And .EOF = False Then
        
            '--- ceci affichera la position de l'enregistrement actif pour ce jeu d'enregistrements ---
            Select Case MemDernierBouton
                Case ETATS_BOUTONS.E_AVANT_NOUVEAU, ETATS_BOUTONS.E_APRES_NOUVEAU
                    Me.Caption = TITRE_FENETRE & " - "
                    LRenseignements.Caption = "-"
                Case Else
                    If IsError(pRecordset("NumCommandeInterne")) = False Then
                        Me.Caption = TITRE_FENETRE & " - " & _
                                         "Commandes internes n° " & pRecordset("NumCommandeInterne") & _
                                         " - Traitement n° " & pRecordset("NumFicheProduction")
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
    
    '--- chargement des détails ---
    If PremiereActivation = True And RechercherSurGrille = False Then
        LectureEnsembleDesDetails
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
    ADODCDetailsChargesProduction.Refresh
    
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

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBLancerRecherche_Click()
    On Error Resume Next
    LanceRechercheOuTri
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

    '--- déclaration ---
    Dim Signet As Variant
    
    If CBRechercherSurGrille.Enabled = True Then

        '--- affectation ---
        RechercherSurGrille = Not (RechercherSurGrille)
                
        '--- affichage ---
        AfficheGrilleRecherche
        
        '--- lancer la lecture des détails ---
        If PremiereActivation = True And RechercherSurGrille = False Then
            LectureEnsembleDesDetails
        End If

    End If

End Sub

Private Sub CBVisualisationGraphesProduction_Click()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer                                                                                        'pour les boucles FOR...NEXT
    Dim NumRedresseur As Integer                                                                'numéro du redresseur
    Dim NumPosteReel As Integer                                                                  'numéro du poste réel
    Dim FicheVideRenseignementsGraphe As RenseignementsGraphe       'fiche vide des renseignements des graphes
    
    '--- RAZ par défaut des renseignements ---
    TRenseignementsGraphe = FicheVideRenseignementsGraphe

    '--- recherche du numéro de redresseur ---
    For a = LBound(TDetailsGammesProduction()) To UBound(TDetailsGammesProduction())
        
        With TDetailsGammesProduction(a)
                    
            If .NumZone <> 0 Then
                    
                If .NomPosteReel <> "" Then
                    
                    '--- affectation du numéro du poste réel ---
                    NumPosteReel = PosteParNom(.NomPosteReel).NumPoste
                    
                    '--- affectation du numéro de redresseur ---
                    Select Case NumPosteReel
                        Case POSTES.P_C13: NumRedresseur = REDRESSEURS.R_C13
                        Case POSTES.P_C14: NumRedresseur = REDRESSEURS.R_C14
                        Case POSTES.P_C15: NumRedresseur = REDRESSEURS.R_C15
                        Case POSTES.P_C16: NumRedresseur = REDRESSEURS.R_C16
                        Case Else
                    End Select
                        
                End If
                
            End If
                
        End With
            
    Next a

    '--- affectation des renseignements ---
    With TRenseignementsGraphe
        .NumFicheProduction = ADODCDetailsChargesProduction.Recordset.Fields("NumFicheProduction")
        .DateEntreeEnLigne = ADODCDetailsChargesProduction.Recordset.Fields("DateEntreeEnLigne")
        .NumRedresseur = NumRedresseur
    End With

    '--- appel de l'écran gérant la visualisation des graphes de production ---
    AppelFenetre F_VISUALISATION_GRAPHES_PRODUCTION

End Sub

Private Sub CTOnglets_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    '--- focus ---
    Select Case CTOnglets.CurrTab
        
        Case ONGLETS.O_DETAILS_CHARGE
            '--- détails de la charge ---
            If MSHFGDetailsChargesProduction.Visible = True Then MSHFGDetailsChargesProduction.SetFocus
        
        Case ONGLETS.O_GAMME_ANODISATION
            '--- gamme Anodisation ---
            If MSHFGDetailsGammesProduction.Visible = True Then MSHFGDetailsGammesProduction.SetFocus
        
        Case ONGLETS.O_DETAILS_FICHE_PRODUCTION
            '--- fiche de production ---
            If MSHFGDetailsFichesProduction.Visible = True Then MSHFGDetailsFichesProduction.SetFocus
        
        Case ONGLETS.O_ALARMES_LIGNE
            '--- alarmes de la ligne ---
            
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
        PremiereActivation = True
        If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
        LectureEnsembleDesDetails
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

Private Sub LNumGamme_Change()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Static NumGammeAnodisation As String

    '--- affichage du nom de la gamme ---
    With LNumGamme
        If .Caption <> "" And .Caption <> NumGammeAnodisation Then
            Bidon = RechercheGammesAnodisation(.Caption)
            LNomGamme.Caption = UN_ESPACE & TTempEnrGammesAnodisation.NomGamme
            NumGammeAnodisation = .Caption
        End If
    End With

End Sub

Private Sub MSHFGDetailsChargesProduction_GotFocus()
    On Error Resume Next
    SFocusTableDetailsChargesProduction.Visible = True
End Sub

Private Sub MSHFGDetailsChargesProduction_LostFocus()
    On Error Resume Next
    SFocusTableDetailsChargesProduction.Visible = False
End Sub

Private Sub MSHFGDetailsFichesProduction_GotFocus()
    On Error Resume Next
    SFocusTableDetailsFichesProduction.Visible = True
End Sub

Private Sub MSHFGDetailsFichesProduction_LostFocus()
    On Error Resume Next
    SFocusTableDetailsFichesProduction.Visible = False
End Sub

Private Sub MSHFGDetailsGammesProduction_GotFocus()
    On Error Resume Next
    SFocusTableDetailsGammesProduction.Visible = True
End Sub

Private Sub MSHFGDetailsGammesProduction_LostFocus()
    On Error Resume Next
    SFocusTableDetailsGammesProduction.Visible = False
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
    ADODCDetailsChargesProduction.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - ADODCDetailsChargesProduction.Width
    LRenseignements.Left = ADODCDetailsChargesProduction.Left
    CBActualiser.Left = ADODCDetailsChargesProduction.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width
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
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondOrange2
    PBBoutons.Picture = ImgFondDesBoutons
        
    '--- divers sur les renseignements ---
    LRenseignements.BackColor = COULEURS.CYAN_0

    '--- gestion des détails ---
    GestionDetailsChargesProduction GG_INITIALISATION
    GestionDetailsGammesProduction GG_INITIALISATION
    GestionDetailsPhasesProduction GG_INITIALISATION
    GestionDetailsFichesProduction GG_INITIALISATION
    
    '--- affectation ---
    CTOnglets.CurrTab = ONGLETS.O_DETAILS_CHARGE
    
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

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFTraçabiliteProduction = Nothing
    
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
        CTOnglets.MousePointer = vbHourglass
    Else
        Me.MousePointer = vbDefault
        CTOnglets.MousePointer = vbDefault
    End If
    
End Sub

Private Sub TBContenant_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri
            If RechercherSurGrille = True Then
                TDBGGrilleRecherche.SetFocus
            Else
            End If
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBContenant_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES
End Sub

Private Sub TBCritereEntre2Valeurs_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sélectionne le texte saisi ---
    With TBCritereEntre2Valeurs(Index)
        If .SelText = "" Then
            .SelStart = 0
            .SelLength = Len(.Text)
        End If
    End With

End Sub

Private Sub TBCritereEntre2Valeurs_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- filtrage des touches ---
    Select Case KeyCode
        
        Case vbKeyReturn, vbKeyDown
                            
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri
            If RechercherSurGrille = True Then TDBGGrilleRecherche.SetFocus
        
        Case Else
            '--- filtrage des touches ---
            FiltreToucheFonction KeyCode, Shift
    
    End Select

End Sub

Private Sub TBCritereEntre2Valeurs_KeyPress(Index As Integer, KeyAscii As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- Filtrage des touches
    FiltreToucheASCII KeyAscii, DONNEES.D_GENERALE_MAJUSCULES

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
                ADODCDetailsChargesProduction.Recordset.MoveFirst
                KeyCode = 0: Shift = 0
            End If
        Case vbKeyEnd
            If Shift = vbCtrlMask Then
                ADODCDetailsChargesProduction.Recordset.MoveLast
                KeyCode = 0: Shift = 0
            End If
        Case vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyPageUp, vbKeyPageDown
        Case vbKeyTab
            If Shift = vbShiftMask Then
                TBContenant.SetFocus
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

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
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
        PBCriteresRecherche.Height = PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Height - PBCriteresRecherche.Top - MARGES.M_BORD_BAS - 13 * Screen.TwipsPerPixelY
        TDBGGrilleRecherche.Visible = True
    End If
    
    '--- hauteur de la grille de recherche ---
    HauteurGrilleRecherche = PBCriteresRecherche.Height - TDBGGrilleRecherche.Top - TDBGGrilleRecherche.Left - 5 * Screen.TwipsPerPixelY
    If HauteurGrilleRecherche > 0 Then
        TDBGGrilleRecherche.Height = HauteurGrilleRecherche
    End If
    
    '--- placer le focus ---
    If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
    
End Sub

Private Sub Form_GotFocus()
    On Error Resume Next
    If TBCommencantPar.Visible = True Then TBCommencantPar.SetFocus
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    CBQuitter_Click
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
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
            '--- au chargement de la fenetre ---
            CBQuitter.Enabled = True
            ADODCDetailsChargesProduction.Enabled = True
            CBActualiser.Enabled = True
            CBQuitter.Enabled = True
            PBCriteresRecherche.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
            ADODCDetailsChargesProduction.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            PBCriteresRecherche.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
            ADODCDetailsChargesProduction.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            PBCriteresRecherche.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            ADODCDetailsChargesProduction.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            PBCriteresRecherche.Enabled = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
            CBQuitter.Enabled = True
            ADODCDetailsChargesProduction.Enabled = False
            CBActualiser.Enabled = False
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
            ADODCDetailsChargesProduction.Enabled = False
            CBActualiser.Enabled = False
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
            If RechercherSurGrille = True Then
                CBRechercherSurGrille_Click
                Me.Refresh
            End If
            ADODCDetailsChargesProduction.Enabled = True
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
            PBCriteresRecherche.Enabled = True
        
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
            If KeyCode = vbKeyReturn Then LanceRechercheOuTri
            If RechercherSurGrille = True Then
                TDBGGrilleRecherche.SetFocus
            Else
            End If
        Case Else
            FiltreToucheFonction KeyCode, Shift
    End Select
End Sub

Private Sub TBCommencantPar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    Select Case Succ(CBRechercherPar.ListIndex)
        Case IDX_RECHERCHER_PAR.IDX_NUM_COMMANDE_INTERNE: FiltreToucheASCII KeyAscii, DONNEES.D_TEXTE_MAJUSCULES_NUMERIQUES, 8        'n° de commande interne
        Case IDX_RECHERCHER_PAR.IDX_NUM_FICHE_PRODUCTION: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 8                                            'n° de la fiche de production
        Case IDX_RECHERCHER_PAR.IDX_DATE_ENTREE_LIGNE: FiltreToucheASCII KeyAscii, DONNEES.D_DATE_JJMMAAAA                                                       'date du GammesRedresseurs
        Case IDX_RECHERCHER_PAR.IDX_CODE_CLIENT: FiltreToucheASCII KeyAscii, DONNEES.D_CODE_CLIENT, 8                                                                     'code client
        Case Else
    End Select
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées :                    TravailSurGrille -> FALSE = Travail sur la fiche
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
    
    '--- commençant par ---
    TBCommencantPar.Text = CommencantPar
    
    '--- contenant ---
    TBContenant.Text = Contenant
    
    '--- initialisation des champs / grilles ---
    GestionGrilleRecherche GG_INITIALISATION
    GestionGrilleRecherche GG_AFFICHAGE
    
    '--- lancement de la recherche ---
    LanceRechercheOuTri

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lance une recherche en fonction des critères
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LanceRechercheOuTri()
    
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
                                              "NumFicheProduction", _
                                              "NumCommandeInterne", _
                                              "DateEntreeEnLigne", _
                                              "CodeClient")

    '--- début de la requête ---
    RequeteSQL = "SELECT DetailsChargesProduction.*, OuiNon.OuiNonTexte AS ChargePrioritaireTexte " & _
                            "FROM DetailsChargesProduction LEFT OUTER JOIN OuiNon ON DetailsChargesProduction.ChargePrioritaire = OuiNon.OuiNonNumerique "
    
    If IdxRecherchePar = IDX_RECHERCHER_PAR.IDX_DATE_ENTREE_LIGNE Then

        
        '--- filtres pour la date ---
        Filtre1 = "(CONVERT(VARCHAR(10), " & RechercherPar & ", 103) LIKE '" & CommencantPar & "%') "
        Filtre2 = "(CONVERT(VARCHAR(10), " & RechercherPar & ", 103) LIKE '%" & Contenant & "%') "
    
    Else
        
        '--- filtres pour chaines de caractères ---
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
    
    '--- fin de la requête ---
    'TODO
    RequeteSQL = RequeteSQL & "ORDER BY "
    Select Case IdxRecherchePar
        Case IDX_RECHERCHER_PAR.IDX_NUM_FICHE_PRODUCTION
            RequeteSQL = RequeteSQL & "CAST(NumFicheProduction AS INT) DESC, CAST((LEFT(NumCommandeInterne,4) + RIGHT(NumCommandeInterne,3)) AS INT) DESC, DateEntreeEnLigne DESC"
        Case IDX_RECHERCHER_PAR.IDX_NUM_COMMANDE_INTERNE
            RequeteSQL = RequeteSQL & "CAST((LEFT(NumCommandeInterne,4) + RIGHT(NumCommandeInterne,3)) AS INT) DESC, DateEntreeEnLigne DESC"
        Case IDX_RECHERCHER_PAR.IDX_DATE_ENTREE_LIGNE
            RequeteSQL = RequeteSQL & "DateEntreeEnLigne DESC, CAST((LEFT(NumCommandeInterne,4) + RIGHT(NumCommandeInterne,3)) AS INT) DESC"
        Case IDX_RECHERCHER_PAR.IDX_CODE_CLIENT
            RequeteSQL = RequeteSQL & "CodeClient, CAST((LEFT(NumCommandeInterne,4) + RIGHT(NumCommandeInterne,3)) AS INT) DESC, DateEntreeEnLigne DESC"
        Case Else
    End Select
    
    '--- debug ---
    'If IdxRecherchePar = IDX_RECHERCHER_PAR.IDX_DATE_ENTREE_LIGNE Then
    '    Debug.Print RequeteSQL
    'End If
    
    With ADODCDetailsChargesProduction
        
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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des détails des charges de la production
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsChargesProduction(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim TypeCouleur As Boolean
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer, _
            PtrLigne As Integer, _
            NbrLignesReferencesClient As Integer
    Dim TempsEnSecondes As Double
    Dim Texte As String
    Dim FicheVide As ImgDetailsChargesProduction, _
            TCopieDetailsChargesProduction(1 To NBR_LIGNES_DETAILS_CHARGES) As ImgDetailsChargesProduction

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des détails ---
            Erase TDetailsChargesProduction()

            '--- initialisation de la grille des détails ---
            With MSHFGDetailsChargesProduction

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_CHARGES + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_CHARGES_PRODUCTION + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 750                                        'épaisseur des titres
                .RowHeightMin = 315
                .Row = 0
                
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_COMMANDE_INTERNE
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Numéro de pointage"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NBR_REPARATIONS
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "R."
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_CODE_CLIENT
                .ColWidth(.Col) = 20 * EPAISSEUR_CARACTERE: .Text = "Code client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NBR_PIECES
                .ColWidth(.Col) = 8 * EPAISSEUR_CARACTERE: .Text = "Nombre de pièces"
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_DESIGNATION
                .ColWidth(.Col) = 50 * EPAISSEUR_CARACTERE: .Text = "Désignation"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_LIGNES_REFERENCES_CLIENT
                .ColWidth(.Col) = 50 * EPAISSEUR_CARACTERE: .Text = "Quantité / référence du client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_MATIERE
                .ColWidth(.Col) = 30 * EPAISSEUR_CARACTERE: .Text = "Matière"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = LBound(TDetailsChargesProduction()) To UBound(TDetailsChargesProduction())
                
                    '--- N° de lignes ---
                    .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_LIGNES
                    .RowHeight(a) = 315                    'épaisseur des lignes
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_COMMANDE_INTERNE
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_CHARGES_PRODUCTION.C_MATIERE
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.ORANGE_1, COULEURS.CYAN_1)
                    
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_COMMANDE_INTERNE

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = LBound(TDetailsChargesProduction()) To UBound(TDetailsChargesProduction())
                TDetailsChargesProduction(a) = FicheVide
            Next a
            With MSHFGDetailsChargesProduction
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données dans le tableau ---
            For a = LBound(TTempEnrDetailsChargesProduction()) To UBound(TTempEnrDetailsChargesProduction())
                With TTempEnrDetailsChargesProduction(a)
                    
                    '--- détails de la charge ---
                    TDetailsChargesProduction(a).NumCommandeInterne = .NumCommandeInterne
                    TDetailsChargesProduction(a).NbrReparations = .NbrReparations
                    TDetailsChargesProduction(a).DateEntreeEnLigne = .DateEntreeEnLigne
                    TDetailsChargesProduction(a).DateArriveeAuDechargement = .DateArriveeAuDechargement
                    TDetailsChargesProduction(a).NumBarre = .NumBarre
                    TDetailsChargesProduction(a).CodeClient = .CodeClient
                    TDetailsChargesProduction(a).NbrPieces = .NbrPieces
                    TDetailsChargesProduction(a).Designation = .Designation
                    TDetailsChargesProduction(a).Matiere = .Matiere
                    
                    '--- gestion des références du client ---
                    'TDetailsChargesProduction(a).NumLignesReferencesClient = ExtraitReferencesClient(.NumCommandeInterne, _
                                                                                                                                                               .NumLignesReferencesClient, _
                                                                                                                                                               NbrLignesReferencesClient)
                    'TDetailsChargesProduction(a).NbrLignesReferencesClient = NbrLignesReferencesClient
                    
                    TDetailsChargesProduction(a).NumGammeAnodisation = .NumGammeAnodisation
                    TDetailsChargesProduction(a).RefGammeAnodisation = .RefGammeAnodisation
                    
                    TDetailsChargesProduction(a).NumFicheProduction = .NumFicheProduction
                    TDetailsChargesProduction(a).ChargePrioritaire = .ChargePrioritaire
                    TDetailsChargesProduction(a).AlarmesLigne = DecodeAlarmesLigne(.AlarmesLigne)
                
                    TDetailsChargesProduction(a).ControleColmatage = .ControleColmatage
                    TDetailsChargesProduction(a).ControleEpaisseurAnodisation = .ControleEpaisseurAnodisation
                    TDetailsChargesProduction(a).ControleColoration = .ControleColoration
                    TDetailsChargesProduction(a).ControleObservations = .ControleObservations
                
                End With
            Next a

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage du numéro de fiche de production ---
            AffichageTexte LNumFicheProduction, UN_ESPACE & TDetailsChargesProduction(1).NumFicheProduction
            
            '--- affichage de la date d'entrée ---
            If TDetailsChargesProduction(1).DateEntreeEnLigne = Empty Then
                Texte = " DATE INCONNUE"
            Else
                Texte = UN_ESPACE & Format(TDetailsChargesProduction(1).DateEntreeEnLigne, FORMAT_DATE_HEURE_1)
            End If
            AffichageTexte LDateEntreeEnLigne, Texte
            
            '--- affichage de la date d'arrivée au déchargement ---
            If TDetailsChargesProduction(1).DateArriveeAuDechargement = Empty Then
                Texte = " DATE INCONNUE"
            Else
                Texte = UN_ESPACE & Format(TDetailsChargesProduction(1).DateArriveeAuDechargement, FORMAT_DATE_HEURE_1)
            End If
            AffichageTexte LDateArriveeAuDechargement, Texte
                            
            '--- affichage du numéro de barre ---
            If TDetailsChargesProduction(1).NumBarre = 0 Then
                Texte = "-"
            Else
                Texte = TDetailsChargesProduction(1).NumBarre
            End If
            AffichageTexte LNumBarre, Texte
            
            '--- affichage indiquant si charge prioritaire ---
            AffichageTexte LChargePrioritaire, UN_ESPACE & IIf(TDetailsChargesProduction(1).ChargePrioritaire > 0, "OUI", "NON")
            
            '--- n° de la gamme d'anodisation ---
            AffichageTexte LNumGamme, TDetailsChargesProduction(1).NumGammeAnodisation
            
            '--- référence de la gamme d'anodisation ---
            AffichageTexte LRefGamme, TDetailsChargesProduction(1).RefGammeAnodisation
            
            '--- matières concernées ---
            If RechercheGammesAnodisation(TDetailsChargesProduction(1).NumGammeAnodisation) = TROUVE Then
                TBMatieresConcernees.Text = TTempEnrGammesAnodisation.TMatieresGamme(1)
                For a = 2 To UBound(TTempEnrGammesAnodisation.TMatieresGamme())
                    If TTempEnrGammesAnodisation.TMatieresGamme(a) <> "" Then
                        TBMatieresConcernees.Text = TBMatieresConcernees.Text & vbCrLf & TTempEnrGammesAnodisation.TMatieresGamme(a)
                    End If
                Next a
            End If
                        
            '--- alarmes de ligne ---
            Texte = TDetailsChargesProduction(1).AlarmesLigne
            If Texte = "" Then
                AffichageTexte TBAlarmesLigne, Texte, COULEURS.BLANC, COULEURS.NOIR
            Else
                AffichageTexte TBAlarmesLigne, Texte, COULEURS.ROUGE_3, COULEURS.JAUNE_3
            End If
            
            '--- construction de la grille ---
            With MSHFGDetailsChargesProduction

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TDetailsChargesProduction()) To UBound(TDetailsChargesProduction())
                    
                    .Row = a
                    If TDetailsChargesProduction(a).NumCommandeInterne = 0 Then
                        TDetailsChargesProduction(a) = FicheVide
                        For b = 1 To NBR_COLONNES_DETAILS_CHARGES_PRODUCTION
                            .Col = b
                            If .Text <> "" Then .Text = ""
                        Next b
                        .RowHeight(a) = .RowHeightMin
                    Else
                        
                        '--- affichage ---
                        Texte = TDetailsChargesProduction(a).NumCommandeInterne
                        AffichageTexteMatrice MSHFGDetailsChargesProduction, a, COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_COMMANDE_INTERNE, Texte
                        
                        Texte = TDetailsChargesProduction(a).CodeClient
                        AffichageTexteMatrice MSHFGDetailsChargesProduction, a, COLONNES_DETAILS_CHARGES_PRODUCTION.C_CODE_CLIENT, Texte
                        
                        Texte = TDetailsChargesProduction(a).NbrPieces
                        AffichageTexteMatrice MSHFGDetailsChargesProduction, a, COLONNES_DETAILS_CHARGES_PRODUCTION.C_NBR_PIECES, Texte
                        
                        Texte = TDetailsChargesProduction(a).Designation
                        AffichageTexteMatrice MSHFGDetailsChargesProduction, a, COLONNES_DETAILS_CHARGES_PRODUCTION.C_DESIGNATION, Texte
                        
                        .Col = COLONNES_DETAILS_CHARGES_PRODUCTION.C_NUM_LIGNES_REFERENCES_CLIENT
                        If .Text <> TDetailsChargesProduction(a).NumLignesReferencesClient Then
                            .Text = TDetailsChargesProduction(a).NumLignesReferencesClient
                            .RowHeight(a) = .RowHeightMin * 0.9 * TDetailsChargesProduction(a).NbrLignesReferencesClient
                        End If
                        
                        'Texte = TDetailsChargesProduction(a).Matiere
                        'AffichageTexteMatrice MSHFGDetailsChargesProduction, a, COLONNES_DETAILS_CHARGES_PRODUCTION.C_MATIERE, Texte
                        
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
' Rôle      : Gestion des détails des fiches de production
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsFichesProduction(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---

    '--- déclaration ---
    Dim TypeCouleur As Boolean
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer
    Dim TempsEnSecondes As Double
    Dim FicheVide As ImgDetailsFichesProduction, _
            TCopieDetailsFichesProduction(1 To NBR_LIGNES_DETAILS_FICHES_PRODUCTION) As ImgDetailsFichesProduction

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des détails ---
            Erase TDetailsFichesProduction()

            '--- initialisation de la grille des détails ---
            With MSHFGDetailsFichesProduction

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_FICHES_PRODUCTION + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_FICHES_PRODUCTION + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 410                                        'épaisseur des titres
                .RowHeightMin = 410
                .Row = 0

                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_NUM_LIGNES
                .ColWidth(.Col) = 4 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_NOM_POSTE
                .ColWidth(.Col) = 9 * EPAISSEUR_CARACTERE: .Text = "Poste"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_TEMPS_REEL_POSTE
                .ColWidth(.Col) = 26 * EPAISSEUR_CARACTERE: .Text = "Temps réel au poste"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_TEMPS_REEL_EGOUTTAGE
                .ColWidth(.Col) = 26 * EPAISSEUR_CARACTERE: .Text = "Temps réel d'égouttage"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_TEMPERATURES
                .ColWidth(.Col) = 18 * EPAISSEUR_CARACTERE: .Text = "Températures"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_REDRESSEUR
                .ColWidth(.Col) = 18 * EPAISSEUR_CARACTERE: .Text = "Redresseurs"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_ANALYSEUR
                .ColWidth(.Col) = 18 * EPAISSEUR_CARACTERE: .Text = "Analyseur"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_ALARMES_POSTE
                .ColWidth(.Col) = 76 * EPAISSEUR_CARACTERE: .Text = "Alarmes de poste"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = LBound(TDetailsFichesProduction()) To UBound(TDetailsFichesProduction())
                
                    '--- N° de lignes ---
                    .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_NUM_LIGNES
                    .RowHeight(a) = 810                    'épaisseur des lignes
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_NOM_POSTE
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_FICHES_PRODUCTION.C_ALARMES_POSTE
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.CYAN_2, COULEURS.JAUNE_2)
                    
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_NOM_POSTE

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = LBound(TDetailsFichesProduction()) To UBound(TDetailsFichesProduction())
                TDetailsFichesProduction(a) = FicheVide
            Next a
            With MSHFGDetailsFichesProduction
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- initialisation du tableau des détails ---
            Erase TDetailsFichesProduction()
            
            '--- transfert des données dans le tableau ---
            For a = LBound(TTempEnrDetailsFichesProduction()) To UBound(TTempEnrDetailsFichesProduction())
                
                With TTempEnrDetailsFichesProduction(a)

                    If .NumPoste >= POSTES.P_CHGT_1 And .NumPoste <= DERNIER_POSTE Then

                        '--- numéro et nom du poste ---
                        TDetailsFichesProduction(a).NomPoste = TEtatsPostes(.NumPoste).DefinitionPoste.NomPoste
                        TDetailsFichesProduction(a).NumPoste = .NumPoste

                        '--- temps réel au poste ---
                        TDetailsFichesProduction(a).TempsReelPoste = "Entrée le " & Format(.DateEntreePoste, FORMAT_DATE_HEURE_1) & vbCr
                        If .DateSortiePoste = Empty Then
                            TDetailsFichesProduction(a).TempsReelPoste = TDetailsFichesProduction(a).TempsReelPoste & "-" & vbCr & "-"
                        Else
                            TempsEnSecondes = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                            TDetailsFichesProduction(a).TempsReelPoste = TDetailsFichesProduction(a).TempsReelPoste & _
                                                                                                           "Sortie le  " & Format(.DateSortiePoste, FORMAT_DATE_HEURE_1) & vbCr & _
                                                                                                           "Temps = " & CTemps2(TempsEnSecondes)
                        End If

                        '--- temps réel d'égouttage ---
                        If .DateDebutEgouttage = Empty Then
                            TDetailsFichesProduction(a).TempsReelEgouttage = "-" & vbCr
                        Else
                            TDetailsFichesProduction(a).TempsReelEgouttage = "Début le " & Format(.DateDebutEgouttage, FORMAT_DATE_HEURE_1) & vbCr
                        End If
                        If .DateFinEgouttage = Empty Then
                            TDetailsFichesProduction(a).TempsReelEgouttage = TDetailsFichesProduction(a).TempsReelEgouttage & "-" & vbCr & "-"
                        Else
                            TempsEnSecondes = DateDiff("s", .DateDebutEgouttage, .DateFinEgouttage)
                            TDetailsFichesProduction(a).TempsReelEgouttage = TDetailsFichesProduction(a).TempsReelEgouttage & _
                                                                                                                 "Fin le  " & Format(.DateFinEgouttage, FORMAT_DATE_HEURE_1) & vbCr & _
                                                                                                                 "Temps = " & CTemps2(TempsEnSecondes)
                        End If

                        '--- températures ---
                        If .TemperatureEnEntree = 0 Then
                            TDetailsFichesProduction(a).Temperatures = "-" & vbCr & "-"
                        Else
                            TDetailsFichesProduction(a).Temperatures = "En entrant : " & Format(.TemperatureEnEntree, FORMAT_TEMPERATURE_1_DECIMALE_UNITE)
                            If .TemperatureEnSortie = 0 Then
                                TDetailsFichesProduction(a).Temperatures = TDetailsFichesProduction(a).Temperatures & vbCr & "-"
                            Else
                                TDetailsFichesProduction(a).Temperatures = TDetailsFichesProduction(a).Temperatures & vbCr & _
                                                                                                         "En sortant : " & Format(.TemperatureEnSortie, FORMAT_TEMPERATURE_1_DECIMALE_UNITE)
                            End If
                        End If

                        '--- redresseur ---
                        If .URedresseur = 0 Then
                            TDetailsFichesProduction(a).Redresseur = "-" & vbCr & "-"
                        Else
                            TDetailsFichesProduction(a).Redresseur = TDetailsFichesProduction(a).Redresseur & vbCr & "U = " & Format(.URedresseur, FORMAT_TENSION_1_DECIMALE_UNITE)
                            If .IRedresseur = 0 Then
                                TDetailsFichesProduction(a).Redresseur = TDetailsFichesProduction(a).Redresseur & vbCr & "-"
                            Else
                                TDetailsFichesProduction(a).Redresseur = TDetailsFichesProduction(a).Redresseur & vbCr & _
                                                                                                      "I = " & Format(.IRedresseur, FORMAT_INTENSITE_ENTIER_UNITE)
                            End If
                        End If

                        '--- analyseur ---
                        If .AnalyseurEnEntree = 0 Then
                            TDetailsFichesProduction(a).Analyseur = "-" & vbCr & "-"
                        Else
                            TDetailsFichesProduction(a).Analyseur = "En entrant : " & Format(.AnalyseurEnEntree, FORMAT_ANALYSEUR_UNITE)
                            If .AnalyseurEnSortie = 0 Then
                                TDetailsFichesProduction(a).Analyseur = TDetailsFichesProduction(a).Analyseur & vbCr & "-"
                            Else
                                TDetailsFichesProduction(a).Analyseur = TDetailsFichesProduction(a).Analyseur & vbCr & _
                                                                                                   "En sortant : " & Format(.AnalyseurEnSortie, FORMAT_ANALYSEUR_UNITE)
                            End If
                        End If

                        '--- alarmes de poste ---
                        TDetailsFichesProduction(a).AlarmesPoste = DecodeAlarmesPoste(.AlarmesPoste)

                    Else

                        Exit For
                    
                    End If

                End With

            Next a

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGDetailsFichesProduction

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TDetailsFichesProduction()) To UBound(TDetailsFichesProduction())
                    .Row = a
                    If TDetailsFichesProduction(a).NumPoste = 0 Then
                        TDetailsFichesProduction(a) = FicheVide
                        For b = 1 To NBR_COLONNES_DETAILS_FICHES_PRODUCTION
                            .Col = b
                            If .Text <> "" Then .Text = ""
                        Next b
                    Else
                        
                        '--- affichage ---
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_NOM_POSTE
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).NomPoste
                        
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_TEMPS_REEL_POSTE
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).TempsReelPoste
                        
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_TEMPS_REEL_EGOUTTAGE
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).TempsReelEgouttage
                        
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_TEMPERATURES
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).Temperatures
                        
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_REDRESSEUR
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).Redresseur
                        
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_ANALYSEUR
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).Analyseur
                        
                        .Col = COLONNES_DETAILS_FICHES_PRODUCTION.C_ALARMES_POSTE
                        AffichageTexte MSHFGDetailsFichesProduction, TDetailsFichesProduction(a).AlarmesPoste
                    
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
' Rôle      : Gestion des détails des gammes d'anodisation de la production
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsGammesProduction(ByVal EtatSouhaite As GESTION_GRILLES)
    
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
    Dim FicheVideDetailsGammesProduction As ImgDetailsGammesProduction, _
            TCopieDetailsgammesProduction(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesProduction

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des détails ---
           Erase TDetailsGammesProduction

            '--- initialisation de la grille des détails ---
            With MSHFGDetailsGammesProduction

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_GAMMES_PRODUCTION + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 750                    'épaisseur des titres
                .RowHeightMin = 315
                .Row = 0

                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Code de la zone"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_LIBELLE_ZONE
                .ColWidth(.Col) = 41.5 * EPAISSEUR_CARACTERE: .Text = "Libellé de la zone"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NOM_POSTE_REEL
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Nom du poste"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                .ColWidth(.Col) = 12 * EPAISSEUR_CARACTERE: .Text = "Temps prévu au POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_DECOMPTE_TEMPS_POSTE_REEL
                .ColWidth(.Col) = 12 * EPAISSEUR_CARACTERE: .Text = "Décompte du temps au POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                .ColWidth(.Col) = 12 * EPAISSEUR_CARACTERE: .Text = "Temps prévu d'égouttage"
                .ColAlignment(.Col) = flexAlignCenterCenter

                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                
                    '--- N° de lignes ---
                    .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NUM_LIGNES
                    .RowHeight(a) = 315                    'épaisseur des lignes
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
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
            For a = LBound(TDetailsGammesProduction()) To UBound(TDetailsGammesProduction())
                TDetailsGammesProduction(a) = FicheVideDetailsGammesProduction
            Next a
            With MSHFGDetailsGammesProduction
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- initialisation du tableau des détails ---
            Erase TDetailsGammesProduction
            
            '--- transfert des données dans le tableau ---
            For a = LBound(TTempEnrDetailsGammesProduction()) To UBound(TTempEnrDetailsGammesProduction())
                With TTempEnrDetailsGammesProduction(a)
                    
                    If .NumZone > 0 Then
                        
                        '--- détails de la gamme ---
                        TDetailsGammesProduction(a).NumZone = .NumZone
                        TDetailsGammesProduction(a).Codezone = TZones(.NumZone).Codezone
                        TDetailsGammesProduction(a).LibelleZone = TZones(.NumZone).LibelleZone
                        If .NumPosteReel >= POSTES.P_CHGT_1 And .NumPosteReel <= DERNIER_POSTE Then
                            TDetailsGammesProduction(a).NomPosteReel = TEtatsPostes(.NumPosteReel).DefinitionPoste.NomPoste
                        End If
                        TDetailsGammesProduction(a).TempsAuPosteTexte = .TempsAuPosteTexte
                        TDetailsGammesProduction(a).TempsEgouttageTexte = .TempsEgouttageTexte
                        TDetailsGammesProduction(a).TempsAuPosteSecondes = .TempsAuPosteSecondes
                        TDetailsGammesProduction(a).TempsEgouttageSecondes = .TempsEgouttageSecondes
                    
                        '--- décompte du temps réel en HH:MM:SS ---
                        If .DecompteDuTempsAuPosteReelSecondes = "" Then
                            TDetailsGammesProduction(a).DecompteDuTempsAuPosteReelTexte = "-"
                        Else
                            If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                TDetailsGammesProduction(a).DecompteDuTempsAuPosteReelTexte = CTemps(CLng(.DecompteDuTempsAuPosteReelSecondes))
                            Else
                                TDetailsGammesProduction(a).DecompteDuTempsAuPosteReelTexte = "-"
                            End If
                        End If
                    
                    End If
                
                End With
            Next a
            
        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---
            PtrLigne = 1
            For a = LBound(TDetailsGammesProduction()) To UBound(TDetailsGammesProduction())
                If TDetailsGammesProduction(a).NumZone <> 0 Then
                    TCopieDetailsgammesProduction(PtrLigne) = TDetailsGammesProduction(a)
                    Inc PtrLigne
                End If
            Next a
            For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                TDetailsGammesProduction(a) = TCopieDetailsgammesProduction(a)
            Next a

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGDetailsGammesProduction

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TDetailsGammesProduction()) To UBound(TDetailsGammesProduction())
                    
                    .Row = a
                    
                    If TDetailsGammesProduction(a).NumZone = 0 Then
                        
                        TDetailsGammesProduction(a) = FicheVideDetailsGammesProduction
                        For b = 1 To NBR_COLONNES_DETAILS_GAMMES_PRODUCTION
                            .Col = b
                            If .Text <> "" Then .Text = ""
                            If .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_DECOMPTE_TEMPS_POSTE_REEL Then
                                .CellPictureAlignment = flexAlignRightTop
                                If .CellPicture <> LoadPicture() Then
                                    Set .CellPicture = LoadPicture()
                                End If
                            End If
                        Next b
                    
                    Else
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_CODE_ZONE
                        AffichageTexte MSHFGDetailsGammesProduction, TDetailsGammesProduction(a).Codezone
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_LIBELLE_ZONE
                        AffichageTexte MSHFGDetailsGammesProduction, TDetailsGammesProduction(a).LibelleZone
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_NOM_POSTE_REEL
                        AffichageTexte MSHFGDetailsGammesProduction, TDetailsGammesProduction(a).NomPosteReel
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_AU_POSTE_TEXTE
                        AffichageTexte MSHFGDetailsGammesProduction, TDetailsGammesProduction(a).TempsAuPosteTexte
                        
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_DECOMPTE_TEMPS_POSTE_REEL
                        AffichageTexte MSHFGDetailsGammesProduction, TDetailsGammesProduction(a).DecompteDuTempsAuPosteReelTexte
                       
                        .Col = COLONNES_DETAILS_GAMMES_PRODUCTION.C_TEMPS_EGOUTTAGE_TEXTE
                        AffichageTexte MSHFGDetailsGammesProduction, TDetailsGammesProduction(a).TempsEgouttageTexte
                        
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
' Rôle      : Gestion des détails des tensions et intensités de la production
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsPhasesProduction(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---

    '--- déclaration ---
    Dim a As Integer

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION, GESTION_GRILLES.GG_VIDAGE
            '--- initialisation du tableau ---
            Erase TDetailsPhasesProduction()
        
            '--- effacement de la zone des redresseurs ---
            FRedresseurs.Visible = False
        
        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- initialisation du tableau des détails ---
            Erase TDetailsPhasesProduction()

            '--- transfert des données dans le tableau ---
            For a = PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4
                With TTempEnrDetailsPhasesProduction(a)

                    If .NumRedresseur >= REDRESSEURS.R_C13 And .NumRedresseur <= REDRESSEURS.R_C16 Then
                        
                        '--- affectation des valeurs ---
                        TDetailsPhasesProduction(a).ModeUouI = .ModeUouI
                        TDetailsPhasesProduction(a).TempsPhase = .TempsPhase
                        TDetailsPhasesProduction(a).UPhase = .UPhase
                        TDetailsPhasesProduction(a).IPhase = .IPhase
                            
                    End If

                End With
            Next a

        Case GESTION_GRILLES.GG_COMPRESSION
        
        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage des valeurs des redresseurs ---
                    
            '--- ne pas afficher la partie redresseur par défaut ---
            FRedresseurs.Visible = False
                        
            '--- interdire les évènements ---
            InterdireEvenements = True
                        
            '--- rendre visible le redresseur se trouvant dans la gamme ---
            For a = LBound(TDetailsGammesProduction()) To UBound(TDetailsGammesProduction())
                
                '--- affichage de la partie redresseur ---
                If TDetailsGammesProduction(a).Codezone = "C13 à C16" Then
                    FRedresseurs.Visible = True
                End If
                
            Next a
            
            '--- affichage des valeurs de programmation pour le redresseur ---
            If FRedresseurs.Visible = True Then
                For a = LBound(TDetailsPhasesProduction()) To UBound(TDetailsPhasesProduction())
                    With TDetailsPhasesProduction(a)
                        LTempsPhases(a).Caption = Right(CTemps2(.TempsPhase), 7)
                        LTensionsPhases(a).Caption = Format(.UPhase, FORMAT_TENSION_1_DECIMALE)
                        LIntensitesPhases(a).Caption = Format(.IPhase, FORMAT_INTENSITE_ENTIER)
                    End With
                Next a
            End If
                        
            '--- calcul du temps total de la gamme redresseur ---
            LTempsTotalGammeRedresseur.Caption = Right(CTemps2(CalculTempsTotalGammeRedresseur()), 7)
                        
            '--- autoriser les évènements ---
            InterdireEvenements = True
                
        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture de l'ensemble des détails
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub LectureEnsembleDesDetails()

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- déclaration ---
    Dim NumFicheProduction As String

    If MemDernierBouton <> ETATS_BOUTONS.E_AVANT_NOUVEAU And _
       MemDernierBouton <> ETATS_BOUTONS.E_APRES_NOUVEAU Then
    
        '--- curseur de la souris ---
        SourisEnAttente True
                            
        '--- vidage des grilles ---
        GestionDetailsChargesProduction GG_VIDAGE
        GestionDetailsChargesProduction GG_AFFICHAGE
        GestionDetailsGammesProduction GG_VIDAGE
        GestionDetailsGammesProduction GG_AFFICHAGE
        GestionDetailsPhasesProduction GG_VIDAGE
        GestionDetailsPhasesProduction GG_AFFICHAGE
        GestionDetailsFichesProduction GG_VIDAGE
        GestionDetailsFichesProduction GG_AFFICHAGE

        With ADODCDetailsChargesProduction.Recordset
        
            If Not .BOF And Not .EOF Then
                
                If .status = adRecOK Then
                
                    If IsNull(.Fields("NumFicheProduction")) = False Then
                
                        '--- affectation ---
                        NumFicheProduction = .Fields("NumFicheProduction")
            
                        '--- recherche ---
                        If RechercheDetailsChargesProduction(NumFicheProduction) = TROUVE Then
                            GestionDetailsChargesProduction GG_TRANSFERT_DONNEES
                            GestionDetailsChargesProduction GG_AFFICHAGE
                        End If
                        If RechercheDetailsGammesProduction(NumFicheProduction) = TROUVE Then
                            GestionDetailsGammesProduction GG_TRANSFERT_DONNEES
                            GestionDetailsGammesProduction GG_AFFICHAGE
                        End If
                        If RechercheDetailsPhasesProduction(NumFicheProduction) = TROUVE Then
                            GestionDetailsPhasesProduction GG_TRANSFERT_DONNEES
                            GestionDetailsPhasesProduction GG_AFFICHAGE
                        End If
                        If RechercheDetailsFichesProduction(NumFicheProduction) = TROUVE Then
                            GestionDetailsFichesProduction GG_TRANSFERT_DONNEES
                            GestionDetailsFichesProduction GG_AFFICHAGE
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
' Rôle      : Marque ou restitue un enregistrement (fonction Bookmark)
' Entrées : MarquageRestitution -> TRUE  = Marquage de l'enregistrement
'                                                       FALSE = Restitution de l'enregistrement
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub MarqueEnregistrement(ByVal MarquageRestitution As Boolean)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Static SignetEnregistrement As Variant

    With ADODCDetailsChargesProduction.Recordset
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
                    .Bold = True                                                                'caractères gras
                End With
                
                .RowHeight = 0                                                               'épaisseur des lignes
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
                
                .AllowRowSelect = True                                                'autoriser la sélection des lignes
                .AllowRowSizing = True                                                'autoriser la modification de l'épaisseur des lignes
                
                
                .DataView = dbgNormalView                                         'présentation normale de la grille
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGGrilleRecherche
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_COMMANDE_INTERNE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Numéro de pointage"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NBR_REPARATIONS)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "R."
                    .Width = EPAISSEUR_CARACTERE * 4
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NUM_FICHE_PRODUCTION)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Ordre de passage"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_DATE_ENTREE_LIGNE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Date d'entrée en ligne"
                    .Width = EPAISSEUR_CARACTERE * 17
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_DATE_ARRIVEE_DECHARGEMENT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Date d'arrivée au déchargement"
                    .Width = EPAISSEUR_CARACTERE * 17
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_CODE_CLIENT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Code client"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_NBR_PIECES)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nombre de pièces"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgRight
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_DESIGNATION)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Désignation"
                    .Width = EPAISSEUR_CARACTERE * 49
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With
                
                With .Columns(COLONNES_GRILLE_RECHERCHE.C_MATIERE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Matière"
                    .Width = EPAISSEUR_CARACTERE * 20
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

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
                    .Width = EPAISSEUR_CARACTERE * 18
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgLeft
                End With

                With .Columns(COLONNES_GRILLE_RECHERCHE.C_CHARGE_PRIORITAIRE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Charge prioritaire"
                    .Width = EPAISSEUR_CARACTERE * 9
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = Me.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

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
    For a = LTempsPhases.LBound To LTempsPhases.UBound
        CalculTempsTotalGammeRedresseur = CalculTempsTotalGammeRedresseur + CTempsTexteEnSecondes(LTempsPhases(a).Caption)
    Next a

End Function


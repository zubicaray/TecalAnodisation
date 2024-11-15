VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{7AA27113-BE6E-4712-824E-EE8CC412FAD9}#1.0#0"; "RT Update Manager.ocx"
Object = "{EBDB8E04-52B4-11D2-A8CF-00105A2E51C1}#1.0#0"; "AppOcxClient.ocx"
Begin VB.MDIForm FPrincipale 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   " TECAL VERBRUGGE - ANODISATION"
   ClientHeight    =   9495
   ClientLeft      =   2010
   ClientTop       =   2730
   ClientWidth     =   14655
   Icon            =   "FPrincipale.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBObjetsTampons 
      Align           =   2  'Align Bottom
      Height          =   990
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   14595
      TabIndex        =   5
      Top             =   7950
      Visible         =   0   'False
      Width           =   14655
      Begin RTUPDATEMANAGERLib.RTUpdateManager RTUpdateManager1 
         Height          =   675
         Left            =   2640
         TabIndex        =   13
         Top             =   120
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   1191
         _StockProps     =   0
         XRadio7E2       =   -1  'True
         XRadio8N1       =   0   'False
         XRadioClipboard =   -1  'True
         XClipboardNew   =   -1  'True
         XRadioFile      =   -1  'True
         XIPPort         =   -1  'True
         XRadioModemPort =   -1  'True
      End
      Begin APPOCXCLIENTLib.AppOcxClient AOCFPrincipale 
         Height          =   720
         Left            =   1215
         TabIndex        =   8
         Top             =   105
         Width           =   1050
         _Version        =   65536
         _ExtentX        =   1852
         _ExtentY        =   1270
         _StockProps     =   0
         NameConfig      =   "FPrincipale"
         PathConfig      =   "C:\Anodisation\Base de données"
      End
      Begin RichTextLib.RichTextBox RTBTampon 
         Height          =   435
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   767
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FPrincipale.frx":08CA
      End
      Begin VB.Label LDonneesTransmisesAPI 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "compteur pour API"
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
         Height          =   375
         Left            =   5940
         TabIndex        =   12
         Top             =   300
         Width           =   2535
      End
   End
   Begin VB.Timer TimerNoyauCentral 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1860
      Top             =   1080
   End
   Begin VB.Timer TimerLigneAlarmes 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2400
      Top             =   1080
   End
   Begin VB.PictureBox PBMessages 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   14595
      TabIndex        =   1
      Top             =   8940
      Width           =   14655
      Begin VB.CommandButton CBAcquittementAlarmes 
         BackColor       =   &H0000FFFF&
         Height          =   375
         Left            =   20520
         MaskColor       =   &H00FF00FF&
         Picture         =   "FPrincipale.frx":0950
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Acquittement des alarmes (coupure du gyrophare et du klaxon) "
         Top             =   60
         UseMaskColor    =   -1  'True
         Width           =   1035
      End
      Begin VB.Label LDateHeureSysteme 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   21600
         TabIndex        =   4
         Top             =   60
         Width           =   6615
      End
      Begin VB.Label LTempsNoyauCentral 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   28260
         TabIndex        =   2
         Top             =   60
         Width           =   435
      End
      Begin VB.Label LMessages 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   20415
      End
   End
   Begin VB.Timer TimerDateHeure 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2940
      Top             =   1080
   End
   Begin MSComDlg.CommonDialog CDDialoguesCommuns 
      Left            =   120
      Top             =   1080
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar COBConteneurOutils 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   1111
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      _CBWidth        =   14655
      _CBHeight       =   630
      _Version        =   "6.7.9816"
      Child1          =   "TOBOutils (0)"
      MinWidth1       =   2295
      MinHeight1      =   570
      Width1          =   2295
      NewRow1         =   0   'False
      Child2          =   "TOBOutils (1)"
      MinWidth2       =   4920
      MinHeight2      =   570
      Width2          =   4920
      NewRow2         =   0   'False
      Child3          =   "TOBOutils (2)"
      MinWidth3       =   1005
      MinHeight3      =   540
      Width3          =   17115
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar TOBOutils 
         Height          =   540
         Index           =   2
         Left            =   7695
         Negotiate       =   -1  'True
         TabIndex        =   11
         Top             =   45
         Width           =   6870
         _ExtentX        =   12118
         _ExtentY        =   953
         ButtonWidth     =   609
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   26
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " "
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TOBOutils 
         Height          =   570
         Index           =   1
         Left            =   2550
         Negotiate       =   -1  'True
         TabIndex        =   10
         Top             =   30
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   1005
         ButtonWidth     =   609
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " "
               Object.Width           =   1e-4
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TOBOutils 
         Height          =   570
         Index           =   0
         Left            =   30
         Negotiate       =   -1  'True
         TabIndex        =   9
         Top             =   30
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1005
         ButtonWidth     =   609
         ButtonHeight    =   953
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   " "
               Object.Width           =   1e-4
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ILGrillesDonnees 
      Left            =   1260
      Top             =   1080
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   16711935
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":0A9A
            Key             =   "fleche noire"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":0CA6
            Key             =   "fleche blanche"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":0EB2
            Key             =   "fleche grise"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":10BE
            Key             =   "fleche rouge"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":12CA
            Key             =   "fleche jaune"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":14D6
            Key             =   "fleche verte"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":16E2
            Key             =   "fleche cyan"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":18EE
            Key             =   "fleche bleue"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":1AFA
            Key             =   "etoile noire"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":1D06
            Key             =   "etoile blanche"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":1F12
            Key             =   "etoile grise"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":211E
            Key             =   "etoile rouge"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":232A
            Key             =   "etoile jaune"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":2536
            Key             =   "etoile verte"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":2742
            Key             =   "etoile cyan"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":294E
            Key             =   "etoile bleue"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":2B5A
            Key             =   "modification noire"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":2D5E
            Key             =   "modification blanche"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":2F62
            Key             =   "modification grise"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":3166
            Key             =   "modification rouge"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":336A
            Key             =   "modification jaune"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":356E
            Key             =   "modification vert"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":3772
            Key             =   "modification cyan"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":3976
            Key             =   "modification bleue"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":3B7A
            Key             =   "indicateur vert"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":3D7E
            Key             =   "indicateur rouge"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ILOutils 
      Left            =   660
      Top             =   1080
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":3F82
            Key             =   "Aide"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":42D4
            Key             =   "Calculatrice"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":4626
            Key             =   "OFChromage"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":4978
            Key             =   "CaracteristiquesLigne"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":4CCA
            Key             =   "ModeOutilsMenuPrincipal"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":501C
            Key             =   "MoteurInference"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":536E
            Key             =   "ModeCyclique"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":56C0
            Key             =   "Manuel"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":5A12
            Key             =   "GammesAnodisation"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":5D64
            Key             =   "Acquittement"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":60B6
            Key             =   "ChargesEnLigne"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":6408
            Key             =   "Acquittement rouge"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":675A
            Key             =   "CyclesPonts"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":6AAC
            Key             =   "ParametresChromage"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":6FFE
            Key             =   "ChargementPrevisionnel"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":7350
            Key             =   "OptionsProgramme"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":76A2
            Key             =   "Cuves"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":79F4
            Key             =   "Defauts"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":7D46
            Key             =   "Regulation"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":8098
            Key             =   "Maintenance"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":83EA
            Key             =   "ApercuAvantImpression"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":873C
            Key             =   "General"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":8A8E
            Key             =   "Redresseurs"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":8DE0
            Key             =   "TracabiliteDeProduction"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":9132
            Key             =   "ProgrammateurCyclique"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":9484
            Key             =   "Equipe"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":97D6
            Key             =   "Annexes"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FPrincipale.frx":9B28
            Key             =   "FinDeJournee"
         EndProperty
      EndProperty
   End
   Begin VB.Menu MenuQuitter 
      Caption         =   "&Quitter"
   End
   Begin VB.Menu MenuDivers 
      Caption         =   "&Divers"
      Begin VB.Menu MenuDiversNettoyageGraphesProduction 
         Caption         =   "&Nettoyage des graphes de production"
      End
      Begin VB.Menu MenuDiversS1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuDiversChargementCheminBDCLIPPER 
         Caption         =   "Recharger le chemin de la base de données CLIPPER après modification dans le fichier configuration.txt"
      End
   End
   Begin VB.Menu MenuFenetres 
      Caption         =   "&Fenêtres"
      WindowList      =   -1  'True
      Begin VB.Menu MenuFenetresActualiser 
         Caption         =   "&Actualiser"
      End
      Begin VB.Menu MenuFenetresS1 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFenetresEnCascade 
         Caption         =   "&En cascade"
      End
      Begin VB.Menu MenuFenetresMosaiqueHorizontale 
         Caption         =   "Mosaique &horizontale"
      End
      Begin VB.Menu MenuFenetresMosaiqueVerticale 
         Caption         =   "Mosaique &verticale"
      End
      Begin VB.Menu MenuFenetresS2 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFenetresMosaiqueCalculee 
         Caption         =   "Mosaique &calculée"
      End
      Begin VB.Menu MenuFenetresS3 
         Caption         =   "-"
      End
      Begin VB.Menu MenuFenetresFermerTout 
         Caption         =   "Fermer tout"
      End
   End
   Begin VB.Menu MenuAPropos 
      Caption         =   "&?"
   End
End
Attribute VB_Name = "FPrincipale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre principale du logiciel
' Nom                    : FPrincipale.frm
' Date de création : 31/07/2000
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit



'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "Fenêtre principale"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

Private Declare Function StrFormatByteSizeW Lib "Shlwapi" ( _
                         ByVal qdw As Currency, _
                         ByVal pSZPuf As Long, _
                         ByVal cchBuf As Long) As Long




Private Sub AOCFPrincipale_EventNewValue(ByVal NbItems As Long, ByVal TabItemName As Variant, ByVal value As Variant, ByVal Quality As Variant, ByVal TimeStamp As Variant)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim ModificationCycleP1 As Boolean             'indique des modifications sur le cycle du pont 1
    Dim ModificationCycleP2 As Boolean             'indique des modifications sur le cycle du pont 2
    Dim ModificationSurChariots As Boolean       'indique des modifications sur les chariots
    Dim a As Integer                                              'pour les boucles FOR...NEXT
    Dim b As Integer                                              'pour les boucles FOR...NEXT
    Dim NumCouche As Integer                             'numéro d'une couche
    Dim NumCharge As Integer                             'numéro de charge
    Dim NumDefaut As Integer                              'numéro d'un défaut
    
    Dim Valeur As Long                                         'représente une valeur quelconque
    Dim Index As Long                                          'représente un index
    
    Static API_PresenceChariots As String * 16, _
              API_VerrouillageChariots As String * 16
    
    Dim DelaisLongEVEau As String * 16             'représente en binaire les délais trop long des électro-vannes
    Dim Cle As String                                            'représente une clé pour une recherche unique
    Dim Texte As String                                         'représente un texte quelconque
    Dim TexteBinaire As String                             'représente un texte avec des 0 ou 1 d'une conversion avec l'automate
    Dim TexteRecherche As String                       'pour la recherche de texte
    
    For a = 1 To NbItems

        'Debug.Print TabItemName(a), Value(a)

        Select Case TabItemName(a)

            '********************************************************************************************************
            '                                                                  ANNEXES
            '********************************************************************************************************
            
            Case "ANODISATION.ANNEXES.MW_Mode_EV_Brillantage"
                '--- mode de l'électro-vanne d'air du brillantage ---
                With TEtatsAnnexes
                    .ModeEVBrillantage = value(a)
                    .API_ChangementsEVBrillantage = True
                End With
            
            Case "ANODISATION.ANNEXES.PeriodiciteEVBrillantage"
                '--- périodicité de l'électro-vanne d'air du brillantage ---
                With TEtatsAnnexes
                    .PeriodiciteEVBrillantage = value(a) / 60
                    .API_ChangementsEVBrillantage = True
                End With
            
            Case "ANODISATION.ANNEXES.TempsMarcheEVBrillantage"
                '--- temps de marche de l'électro-vanne d'air du brillantage ---
                With TEtatsAnnexes
                    .TempsMarcheEVBrillantage = value(a)
                    .API_ChangementsEVBrillantage = True
                End With
            
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            Case "ANODISATION.ANNEXES.MW_Mode_EV_Eau_Ligne"
                '--- mode de l'électro-vanne d'arrivée d'eau de la ligne ---
                With TEtatsAnnexes
                    .ModeEVEauLigne = value(a)
                    .API_ChangementsEVEauLigne = True
                End With
            
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            Case "ANODISATION.ANNEXES.MW_Mode_Compres_P1"
                '--- mode du compresseur du pont 1 ---
                With TEtatsAnnexes
                    .ModeCompresseurP1 = value(a)
                    .API_ChangementsCompresseurP1 = True
                End With
            
            Case "ANODISATION.ANNEXES.A_Compresseur_P1"
                '--- états du compresseur du pont 1 ---
                With TEtatsAnnexes
                    .EtatsCompresseurP1 = IIf(value(a) = True, ETATS_COMPRESSEURS_PONTS.E_MARCHE, ETATS_COMPRESSEURS_PONTS.E_ARRET)
                    .API_ChangementsCompresseurP1 = True
                End With
            
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            Case "ANODISATION.ANNEXES.MW_Mode_Compres_P2"
                '--- mode du compresseur du pont 2 ---
                With TEtatsAnnexes
                    .ModeCompresseurP2 = value(a)
                    .API_ChangementsCompresseurP2 = True
                End With
            
            Case "ANODISATION.ANNEXES.A_Compresseur_P2"
                '--- états du compresseur du pont 2 ---
                With TEtatsAnnexes
                    .EtatsCompresseurP2 = IIf(value(a) = True, ETATS_COMPRESSEURS_PONTS.E_MARCHE, ETATS_COMPRESSEURS_PONTS.E_ARRET)
                    .API_ChangementsCompresseurP2 = True
                End With
            
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            Case "ANODISATION.ANNEXES.MW_Mode_Eclairage_P1"
                '--- mode de l'éclairage du pont 1 ---
                With TEtatsAnnexes
                    .ModeEclairageP1 = value(a)
                    .API_ChangementsEclairageP1 = True
                End With
            
            Case "ANODISATION.ANNEXES.A_Eclairage_P1"
                '--- états de l'éclairage du pont 1 ---
                With TEtatsAnnexes
                    .EtatsEclairageP1 = IIf(value(a) = True, ETATS_ECLAIRAGE_PONTS.E_MARCHE, ETATS_ECLAIRAGE_PONTS.E_ARRET)
                    .API_ChangementsEclairageP1 = True
                End With
            
            '------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            Case "ANODISATION.ANNEXES.A_Eclairage_P2"
                '--- états de l'éclairage du pont 2 ---
                With TEtatsAnnexes
                    .EtatsEclairageP2 = IIf(value(a) = True, ETATS_ECLAIRAGE_PONTS.E_MARCHE, ETATS_ECLAIRAGE_PONTS.E_ARRET)
                    .API_ChangementsEclairageP2 = True
                End With
            
            Case "ANODISATION.ANNEXES.MW_Mode_Eclairage_P2"
                '--- mode de l'éclairage du pont 2 ---
                With TEtatsAnnexes
                    .ModeEclairageP2 = value(a)
                    .API_ChangementsEclairageP2 = True
                End With
            
            '********************************************************************************************************
            '                                                              ETATS DE LA LIGNE
            '********************************************************************************************************

            Case "ANODISATION.ETATS_LIGNE.EtatsCommutations"
                '--- de maintenance à automatique pont 1 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                If Bit(TexteBinaire, POS_BIT_MAINTENANCE_P1) = 1 Then
                    TEtatsPonts(PONTS.P_1).ModePont = MODES_PONTS.M_MAINTENANCE
                Else
                    If Bit(TexteBinaire, POS_BIT_MANUEL_P1) = 1 Then
                        TEtatsPonts(PONTS.P_1).ModePont = MODES_PONTS.M_MANUEL
                    ElseIf Bit(TexteBinaire, POS_BIT_SEMI_AUTOMATIQUE_P1) = 1 Then
                        TEtatsPonts(PONTS.P_1).ModePont = MODES_PONTS.M_SEMI_AUTOMATIQUE
                    ElseIf Bit(TexteBinaire, POS_BIT_AUTOMATIQUE_P1) = 1 Then
                        TEtatsPonts(PONTS.P_1).ModePont = MODES_PONTS.M_AUTOMATIQUE
                    End If
                End If

                '--- de maintenance à automatique pont 2 ---
                If Bit(TexteBinaire, POS_BIT_MAINTENANCE_P2) = 1 Then
                    TEtatsPonts(PONTS.P_2).ModePont = MODES_PONTS.M_MAINTENANCE
                Else
                    If Bit(TexteBinaire, POS_BIT_MANUEL_P2) = 1 Then
                        TEtatsPonts(PONTS.P_2).ModePont = MODES_PONTS.M_MANUEL
                    ElseIf Bit(TexteBinaire, POS_BIT_SEMI_AUTOMATIQUE_P2) = 1 Then
                        TEtatsPonts(PONTS.P_2).ModePont = MODES_PONTS.M_SEMI_AUTOMATIQUE
                    ElseIf Bit(TexteBinaire, POS_BIT_AUTOMATIQUE_P2) = 1 Then
                        TEtatsPonts(PONTS.P_2).ModePont = MODES_PONTS.M_AUTOMATIQUE
                    End If
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.EtatsSecuritesLigne"
                '--- sécurités de la ligne ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                
                TEtatsLigne.MarcheGenerale = Bit(TexteBinaire, POS_BIT_MARCHE_GENERALE)
                TEtatsLigne.ArretGeneral = Not (TEtatsLigne.MarcheGenerale)
                
                TEtatsLigne.StopLigne = Bit(TexteBinaire, POS_BIT_STOP_LIGNE)
                TEtatsLigne.ArretUrgenceP1 = Bit(TexteBinaire, POS_BIT_ARRET_URGENCE_P1)
                TEtatsLigne.ArretUrgenceP2 = Bit(TexteBinaire, POS_BIT_ARRET_URGENCE_P2)
                
                TEtatsLigne.ArretUrgence = Bit(TexteBinaire, POS_BIT_ARRET_URGENCE)
                TEtatsLigne.PortillonsLigneVie = Bit(TexteBinaire, POS_BIT_PORTILLONS_LIGNE_VIE)
                TEtatsLigne.SecuriteP1 = Bit(TexteBinaire, POS_BIT_SECURITE_P1)
                TEtatsLigne.SecuriteP2 = Bit(TexteBinaire, POS_BIT_SECURITE_P2)
                TEtatsLigne.ManqueTension = Bit(TexteBinaire, POS_BIT_MANQUE_TENSION)
                TEtatsLigne.ManqueAir = Bit(TexteBinaire, POS_BIT_MANQUE_AIR)
                TEtatsLigne.AcquittementsDefauts = Bit(TexteBinaire, POS_BIT_ACQUITTEMENT_DEFAUTS)
                TEtatsLigne.FrontMontantDefauts = Bit(TexteBinaire, POS_BIT_FRONT_MONTANT_DEFAUTS)

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.EtatsPresenceChariots"
                '--- présence chariots des chariots chargement / déchargement ---
                API_PresenceChariots = CBin(Val("&h" & CHex(value(a))))
                ModificationSurChariots = True

            Case "ANODISATION.ETATS_LIGNE.EtatsVerrouChariots"
                '--- verrouillage chariots des chariots chargement / déchargement ---
                API_VerrouillageChariots = CBin(Val("&h" & CHex(value(a))))
                ModificationSurChariots = True

            '********************************************************************************************************
            '                                                        ETATS DE LA LIGNE PONT 1
            '********************************************************************************************************

            Case "ANODISATION.ETATS_LIGNE.EntreesP1"
                '--- entrées du pont 1 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_1).TEntreesAPI
                    .E_NiveauBas = Bit(TexteBinaire, BIT_0)
                    .E_NiveauIntermediaire = Bit(TexteBinaire, BIT_1)
                    .E_NiveauHaut = Bit(TexteBinaire, BIT_2)
                    .M_AccrochesEnHaut = Bit(TexteBinaire, BIT_3)
                    .M_AccrochesEnBas = Bit(TexteBinaire, BIT_4)
                    
                    .M_MoteurTourneTrlPont = Bit(TexteBinaire, BIT_8)
                    .M_MoteurTourneLevPont = Bit(TexteBinaire, BIT_9)
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.SortiesP1"
                '--- sorties du pont 1 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_1).TSortiesAPI
                    .S_EVMonteeAccroches = Bit(TexteBinaire, BIT_0)
                    .S_EVDescenteAccroches = Bit(TexteBinaire, BIT_1)
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.DefautsP1"
                '--- défauts du pont 1 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_1).TEntreesAPI
                    
                    .M_DefautVariateurTrlPont = Bit(TexteBinaire, BIT_0)
                    '.M_AxeNonReferenceTrlPont = Not (Bit(TexteBinaire, BIT_1))
                    .M_SurcourseTrlAvant = Bit(TexteBinaire, BIT_2)
                    .M_SurcourseTrlArriere = Bit(TexteBinaire, BIT_3)
                    
                    .M_DefautVariateurLevPont = Bit(TexteBinaire, BIT_4)
                    .M_AxeNonReferenceLevPont = IIf(Bit(TexteBinaire, BIT_5) = 0, True, False)          'ATTENTION inversion du bit
                    .M_SurcourseLevHaut = Bit(TexteBinaire, BIT_6)
                    .M_SurcourseLevBas = Bit(TexteBinaire, BIT_7)

                    .M_DelaiTropLongDescenteAccroches = Bit(TexteBinaire, BIT_8)
                    .M_DelaiTropLongMonteeAccroches = Bit(TexteBinaire, BIT_9)
                
                    .M_DefautPresencePicece = Bit(TexteBinaire, BIT_10)
                
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.AxesP1"
                '--- axes du pont 1 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_1).TEntreesAPI
                    .M_MarquageAxeTrL = Bit(TexteBinaire, BIT_0)
                    .M_MarquagePVTrL = Bit(TexteBinaire, BIT_1)
                    .M_MarquageMVTrL = Bit(TexteBinaire, BIT_2)
                    .M_MarquageArretTrL = Bit(TexteBinaire, BIT_3)
                    .M_MarquageAxeLev = Bit(TexteBinaire, BIT_4)
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.PosteActuelP1"
                '--- poste actuel du pont 1 ---
                Valeur = value(a)
                If Valeur >= POSTES.P_CHGT_1 And Valeur <= DERNIER_POSTE Then
                    With TEtatsPonts(PONTS.P_1)
                        
                        '--- poste actuel ---
                        .PosteActuel = Valeur
                
                        '--- sens ---
                        If .PosteActuel = .PosteDestination Then
                            .SensX = SENS_X.S_AU_POSTE
                        ElseIf .PosteActuel < .PosteDestination Then
                            .SensX = SENS_X.S_AVANT
                        ElseIf .PosteActuel > .PosteDestination Then
                            .SensX = SENS_X.S_ARRIERE
                        End If
                
                    End With
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.PosteDestinationP1"
                '--- poste de destination du pont 1 ---
                Valeur = value(a)
                If Valeur >= POSTES.P_CHGT_1 And Valeur <= DERNIER_POSTE Then
                    With TEtatsPonts(PONTS.P_1)
                        
                        '--- poste de destination ---
                        .PosteDestination = Valeur
                
                        '--- sens ---
                        If .PosteActuel = .PosteDestination Then
                            .SensX = SENS_X.S_AU_POSTE
                        ElseIf .PosteActuel < .PosteDestination Then
                            .SensX = SENS_X.S_AVANT
                        ElseIf .PosteActuel > .PosteDestination Then
                            .SensX = SENS_X.S_ARRIERE
                        End If
                
                    End With
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.NiveauActuelP1"
                '--- niveau actuel du pont 1 ---
                Valeur = value(a)
                If Valeur >= NIVEAUX_PONTS.N_BAS And Valeur <= NIVEAUX_PONTS.N_HAUT Then
                    TEtatsPonts(PONTS.P_1).NiveauActuel = Valeur
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.NiveauDestinationP1"
                '--- niveau de destination du pont 1 ---
                Valeur = value(a)
                If Valeur >= NIVEAUX_PONTS.N_BAS And Valeur = NIVEAUX_PONTS.N_HAUT Then
                    TEtatsPonts(PONTS.P_1).NiveauDestination = Valeur
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.PointeurActionP1"
                '--- pointeur de l'action du pont 1 ---
                TEtatsPonts(PONTS.P_1).PtrEtActionEnCoursAPI.PtrAction = value(a)

            Case "ANODISATION.ETATS_LIGNE.NumActionP1"
                '--- numéro de l'action du pont 1 ---
                TEtatsPonts(PONTS.P_1).PtrEtActionEnCoursAPI.NumAction = value(a)

            '********************************************************************************************************
            '                                                        ETATS DE LA LIGNE PONT 2
            '********************************************************************************************************

            Case "ANODISATION.ETATS_LIGNE.EntreesP2"
                '--- entrées du pont 2 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_2).TEntreesAPI
                    .E_NiveauBas = Bit(TexteBinaire, BIT_0)
                    .E_NiveauIntermediaire = Bit(TexteBinaire, BIT_1)
                    .E_NiveauHaut = Bit(TexteBinaire, BIT_2)
                    .M_AccrochesEnHaut = Bit(TexteBinaire, BIT_3)
                    .M_AccrochesEnBas = Bit(TexteBinaire, BIT_4)
                    
                    .M_MoteurTourneTrlPont = Bit(TexteBinaire, BIT_8)
                    .M_MoteurTourneLevPont = Bit(TexteBinaire, BIT_9)
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.SortiesP2"
                '--- sorties du pont 2 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_2).TSortiesAPI
                    .S_EVMonteeAccroches = Bit(TexteBinaire, BIT_0)
                    .S_EVDescenteAccroches = Bit(TexteBinaire, BIT_1)
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.DefautsP2"
                '--- défauts du pont 2 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_2).TEntreesAPI

                    .M_DefautVariateurTrlPont = Bit(TexteBinaire, BIT_0)
                    '.M_AxeNonReferenceTrlPont = Not (Bit(TexteBinaire, BIT_1))
                    .M_SurcourseTrlAvant = Bit(TexteBinaire, BIT_2)
                    .M_SurcourseTrlArriere = Bit(TexteBinaire, BIT_3)
                    
                    .M_DefautVariateurLevPont = Bit(TexteBinaire, BIT_4)
                    .M_AxeNonReferenceLevPont = IIf(Bit(TexteBinaire, BIT_5) = 0, True, False)          'ATTENTION inversion du bit
                    .M_SurcourseLevHaut = Bit(TexteBinaire, BIT_6)
                    .M_SurcourseLevBas = Bit(TexteBinaire, BIT_7)

                    .M_DelaiTropLongDescenteAccroches = Bit(TexteBinaire, BIT_8)
                    .M_DelaiTropLongMonteeAccroches = Bit(TexteBinaire, BIT_9)

                    .M_DefautPresencePicece = Bit(TexteBinaire, BIT_10)

                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.AxesP2"
                '--- axes du pont 2 ---
                TexteBinaire = CBin(Val("&h" & CHex(value(a))))
                With TEtatsPonts(PONTS.P_2).TEntreesAPI
                    .M_MarquageAxeTrL = Bit(TexteBinaire, BIT_0)
                    .M_MarquagePVTrL = Bit(TexteBinaire, BIT_1)
                    .M_MarquageMVTrL = Bit(TexteBinaire, BIT_2)
                    .M_MarquageArretTrL = Bit(TexteBinaire, BIT_3)
                    .M_MarquageAxeLev = Bit(TexteBinaire, BIT_4)
                End With

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.PosteActuelP2"
                '--- poste actuel du pont 2 ---
                Valeur = value(a)
                If Valeur >= POSTES.P_CHGT_1 And Valeur <= DERNIER_POSTE Then
                    With TEtatsPonts(PONTS.P_2)
                
                        'Log ("Event new value:" + .PosteDestination)
                    
                        '--- poste actuel ---
                        .PosteActuel = Valeur
                
                        '--- sens ---
                        If .PosteActuel = .PosteDestination Then
                            .SensX = SENS_X.S_AU_POSTE
                        ElseIf .PosteActuel < .PosteDestination Then
                            .SensX = SENS_X.S_AVANT
                        ElseIf .PosteActuel > .PosteDestination Then
                            .SensX = SENS_X.S_ARRIERE
                        End If
                
                    End With
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.PosteDestinationP2"
                '--- poste de destination du pont 2 ---
                Valeur = value(a)
                If Valeur >= POSTES.P_CHGT_1 And Valeur <= DERNIER_POSTE Then
                    With TEtatsPonts(PONTS.P_2)
                        
                        '--- poste de destination ---
                        .PosteDestination = Valeur
                
                        '--- sens ---
                        If .PosteActuel = .PosteDestination Then
                            .SensX = SENS_X.S_AU_POSTE
                        ElseIf .PosteActuel < .PosteDestination Then
                            .SensX = SENS_X.S_AVANT
                        ElseIf .PosteActuel > .PosteDestination Then
                            .SensX = SENS_X.S_ARRIERE
                        End If
                
                    End With
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.NiveauActuelP2"
                '--- niveau actuel du pont 2 ---
                Valeur = value(a)
                If Valeur >= NIVEAUX_PONTS.N_BAS And Valeur <= NIVEAUX_PONTS.N_HAUT Then
                    TEtatsPonts(PONTS.P_2).NiveauActuel = Valeur
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.NiveauDestinationP2"
                '--- niveau de destination du pont 2 ---
                Valeur = value(a)
                If Valeur >= NIVEAUX_PONTS.N_BAS And Valeur = NIVEAUX_PONTS.N_HAUT Then
                    TEtatsPonts(PONTS.P_2).NiveauDestination = Valeur
                End If

            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.ETATS_LIGNE.PointeurActionP2"
                '--- pointeur de l'action du pont 2 ---
                TEtatsPonts(PONTS.P_2).PtrEtActionEnCoursAPI.PtrAction = value(a)

            Case "ANODISATION.ETATS_LIGNE.NumActionP2"
                '--- numéro de l'action du pont 2 ---
                TEtatsPonts(PONTS.P_2).PtrEtActionEnCoursAPI.NumAction = value(a)

            '********************************************************************************************************
            '                                               LASERS et CODEURS des LEVAGES
            '********************************************************************************************************

            Case "ANODISATION.LASERS_CODEURS_LEV.MD_Mesure_Laser_P1"
                '--- laser actuel du pont 1 ---
                TEtatsPonts(PONTS.P_1).PositionActuelleLaserTrlPont = value(a)
            
            Case "ANODISATION.LASERS_CODEURS_LEV.MD_Destin_Laser_P1"
                '--- laser de destination du pont 1 ---
                TEtatsPonts(PONTS.P_1).PositionCibleLaserTrlPont = value(a)
            
            Case "ANODISATION.LASERS_CODEURS_LEV.MD_Actuelle_Cod_Lev_P1"
                '--- valeur actuelle du levage du pont 1 ---
                TEtatsPonts(PONTS.P_1).PositionActuelleCodeurLevPont = value(a)
            
            '------------------------------------------------------------------------------------------------------------------------------------------------------------

            Case "ANODISATION.LASERS_CODEURS_LEV.MD_Mesure_Laser_P2"
                '--- laser actuel du pont 2 ---
                TEtatsPonts(PONTS.P_2).PositionActuelleLaserTrlPont = value(a)
            
            Case "ANODISATION.LASERS_CODEURS_LEV.MD_Destin_Laser_P2"
                '--- laser de destination du pont 2 ---
                TEtatsPonts(PONTS.P_2).PositionCibleLaserTrlPont = value(a)
            
            Case "ANODISATION.LASERS_CODEURS_LEV.MD_Actuelle_Cod_Lev_P2"
                '--- valeur actuelle du levage du pont 2 ---
                TEtatsPonts(PONTS.P_2).PositionActuelleCodeurLevPont = value(a)

            '********************************************************************************************************
            '                                                              POIDS SOULEVES
            '********************************************************************************************************
            
            Case "ANODISATION.POIDS_SOULEVES.MD_Poids_Lev_P1"
                '--- poids soulevés du PONT 1 ---
                TEtatsPonts(PONTS.P_1).PoidsSouleve = value(a)
            
            Case "ANODISATION.POIDS_SOULEVES.MD_Poids_Lev_P2"
                '--- poids soulevés du PONT 2 ---
                TEtatsPonts(PONTS.P_2).PoidsSouleve = value(a)
            
            '********************************************************************************************************
            '                                                              SUIVI DES PONTS
            '********************************************************************************************************

            Case "ANODISATION.SUIVI_LIGNE.NumChargeP1"
                TEtatsPonts(PONTS.P_1).NumCharge = value(a)
            
            Case "ANODISATION.SUIVI_LIGNE.OptionsGamme1P1"
                TEtatsPonts(PONTS.P_1).OptionsGamme1 = value(a)
            
            Case "ANODISATION.SUIVI_LIGNE.OptionsGamme2P1"
                TEtatsPonts(PONTS.P_1).OptionsGamme2 = value(a)
            
            Case "ANODISATION.SUIVI_LIGNE.NumChargeP2"
                TEtatsPonts(PONTS.P_2).NumCharge = value(a)

            Case "ANODISATION.SUIVI_LIGNE.OptionsGamme1P2"
                TEtatsPonts(PONTS.P_2).OptionsGamme1 = value(a)
            
            Case "ANODISATION.SUIVI_LIGNE.OptionsGamme2P2"
                TEtatsPonts(PONTS.P_2).OptionsGamme2 = value(a)

            Case Else
                '********************************************************************************************************
                '                                                           ACTIONS DU PONT 1
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ACTIONS_P1.Action") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    TImageAPICyclesPonts(PONTS.P_1, Index) = value(a)
                    ModificationCycleP1 = True
                End If

                '********************************************************************************************************
                '                                                           ACTIONS DU PONT 2
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ACTIONS_P2.Action") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    TImageAPICyclesPonts(PONTS.P_2, Index) = value(a)
                    ModificationCycleP2 = True
                End If

                '********************************************************************************************************
                '                                                            SUIVI DE LA LIGNE
                '********************************************************************************************************
                If InStr(TabItemName(a), ".SUIVI_LIGNE.NumChargePoste") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    TEtatsPostes(Index).NumCharge = value(a)
                End If

                If InStr(TabItemName(a), ".SUIVI_LIGNE.CondamnationPoste") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    TEtatsPostes(Index).Condamnation = IIf(value(a) = 0, False, True)
                End If

                '********************************************************************************************************
                '                                                                 CHARGES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".CHARGE_") > 0 Then
                            
                    '--- extraction du numéro de charge ---
                    For b = CHARGES.C_NUM_MINI To CHARGES.C_NUM_MAXI
                        Cle = "CHARGE_" & Right("00" & b, 2) & "."                       'affectation de la clé
                        If InStr(TabItemName(a), Cle) > 0 Then                               'recherche de la clé
                            NumCharge = b                                                               'affectation du numéro de charge
                            Exit For                                                                            'sortie si recherche positive
                        End If
                    Next b
                    
                    '--- effectuer l'analyse uniquement avec un numéro de charge supérieur à 0 ---
                    If NumCharge > 0 Then

                        With TEtatsCharges(NumCharge)

                            '--- numéro de commande interne ---
                            If InStr(TabItemName(a), Cle & "NumCommandeInterne") > 0 Then
                                Texte = Right(String(7, "0") & CStr(value(a)), 7)
                                .TDetailsCharges(1).NumCommandeInterne = Left(Texte, 4) & "C" & Right(Texte, 3)
                            End If
                            
                            '--- le numéro de barre ---
                            If InStr(TabItemName(a), Cle & "NumBarre") > 0 Then .Numbarre = value(a)
                            
                            '--- mode de travail du redresseur U(tension)=0, I(intensité)=1 ---
                            If InStr(TabItemName(a), Cle & "ModeUouI") > 0 Then .ModeUouI = value(a)
                            
                            '--- phase 1 complète ---
                            If InStr(TabItemName(a), Cle & "TpsPhase1") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1).TempsPhase = value(a)
                            If InStr(TabItemName(a), Cle & "UPhase1") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1).UPhase = value(a) / 10
                            If InStr(TabItemName(a), Cle & "IPhase1") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1).IPhase = value(a)
                            
                            '--- phase 2 complète ---
                            If InStr(TabItemName(a), Cle & "TpsPhase2") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T2).TempsPhase = value(a)
                            If InStr(TabItemName(a), Cle & "UPhase2") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T2).UPhase = value(a) / 10
                            If InStr(TabItemName(a), Cle & "IPhase2") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T2).IPhase = value(a)

                            '--- phase 3 complète ---
                            If InStr(TabItemName(a), Cle & "TpsPhase3") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T3).TempsPhase = value(a)
                            If InStr(TabItemName(a), Cle & "UPhase3") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T3).UPhase = value(a) / 10
                            If InStr(TabItemName(a), Cle & "IPhase3") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T3).IPhase = value(a)
                            
                            '--- phase 4 complète ---
                            If InStr(TabItemName(a), Cle & "TpsPhase4") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase = value(a)
                            If InStr(TabItemName(a), Cle & "UPhase4") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).UPhase = value(a) / 10
                            If InStr(TabItemName(a), Cle & "IPhase4") > 0 Then .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).IPhase = value(a)

                        End With
                    
                    End If
                
                End If
                
                '********************************************************************************************************
                '                                                               REDRESSEURS
                '********************************************************************************************************
                If InStr(TabItemName(a), ".REDRESSEURS.") > 0 Then
                    
                    '--- analyse du dernier caractère pour trouver l'index du redresseur ---
                    Texte = Right(TabItemName(a), 1)
                    
                    If IsNumeric(Texte) = True Then
                    
                        '--- affectation de l'index ---
                        Index = CLng(Texte)
    
                        With TEtatsRedresseurs(Index)
    
                            '--- marche / arrêt et  manu / auto ---
                            If InStr(TabItemName(a), "EtatsMarcheManuAutoR") > 0 Then
    
                                '--- conversion en binaire ---
                                .EtatsMarcheArret = CBin(value(a))
                                
                                '--- affectation des modes du redresseurs ---
                                If .RetoursVersPC = CODE_EXCLUSION_REDRESSEUR Then
                                    
                                    .EtatRedresseur = ETATS_REDRESSEUR.ER_EXCLUSION
                                
                                Else
                                
                                    Select Case value(a)
                                        
                                        Case 0
                                            '--- redresseur en arrêt ---
                                            .EtatRedresseur = ETATS_REDRESSEUR.ER_ARRET
                                            .ModeRedresseur = MODES_REDRESSEUR.MR_NON_DEFINI
                                        
                                        Case 1
                                            '--- redresseur en marche et en manuel ---
                                            .EtatRedresseur = ETATS_REDRESSEUR.ER_MARCHE
                                            .ModeRedresseur = MODES_REDRESSEUR.MR_MANUEL
                                        
                                        Case 3
                                            '--- redresseur en marche et en automatique ---
                                            .EtatRedresseur = ETATS_REDRESSEUR.ER_MARCHE
                                            .ModeRedresseur = MODES_REDRESSEUR.MR_AUTOMATIQUE
                                    
                                        Case Else
                                    End Select
                                
                                End If
                                
                            End If
                            
                            '--- états 1 du redresseur ---
                            If InStr(TabItemName(a), "Etats1R") > 0 Then
    
                                '--- affectation du mot 1 des états ---
                                .Etats1 = CBin(value(a))
    
                                '--- analyse en fonction des bits ---
                                '.TEntreesAPI.M_DefautRedresseur1DuGroupe = (Bit(.Etats1, 0) = False Or Bit(.Etats1, 1) = True Or Bit(.Etats1, 2) = True Or Bit(.Etats1, 3) = True)
                                '.TEntreesAPI.M_DefautRedresseur2DuGroupe = (Bit(.Etats1, 8) = False Or Bit(.Etats1, 9) = True Or Bit(.Etats1, 10) = True Or Bit(.Etats1, 11) = True)
    
                            End If
    
                            '--- états 2 du redresseur ---
                            If InStr(TabItemName(a), "Etats2R") > 0 Then
    
                                '--- affectation du mot 2 des états ---
                                .Etats2 = CBin(value(a))
    
                                '--- analyse en fonction des bits ---
                                '.TEntreesAPI.M_DefautRedresseur3DuGroupe = (Bit(.Etats2, 0) = False Or Bit(.Etats2, 1) = True Or Bit(.Etats2, 2) = True Or Bit(.Etats2, 3) = True)
                                '.TEntreesAPI.M_DefautRedresseur4DuGroupe = (Bit(.Etats2, 8) = False Or Bit(.Etats2, 9) = True Or Bit(.Etats2, 10) = True Or Bit(.Etats2, 11) = True)
    
                            End If
    
                            '--- demandes du PC ---
                            If InStr(TabItemName(a), "DemandesDuPCR") > 0 Then
                                .DemandesDuPC = value(a)
                            End If
    
                            '--- retours vers le PC ---
                            If InStr(TabItemName(a), "RetoursVersPCR") > 0 Then
                                .RetoursVersPC = value(a)
                                If .RetoursVersPC = CODE_EXCLUSION_REDRESSEUR Then
                                    .EtatRedresseur = ETATS_REDRESSEUR.ER_EXCLUSION
                                End If
                            End If
    
                            '--- mesure de la tension ---
                            If InStr(TabItemName(a), "UMesureR") > 0 Then
                                .U = value(a) / 10
                            End If
                            
                            '--- consigne de la tension ---
                            If InStr(TabItemName(a), "UDemandePCR") > 0 Then
                                .ConsigneU = value(a) / 10
                            End If
                            
                            '--- mesure de l'intensité ---
                            If InStr(TabItemName(a), "IMesureR") > 0 Then
                                .I = value(a)
                            End If
                            
                            '--- consigne de l'intensité ---
                            If InStr(TabItemName(a), "IDemandePCR") > 0 Then
                                .ConsigneI = value(a)
                            End If
    
                            '--- début du cycle ---
                            If InStr(TabItemName(a), "M_Debut_Cycle_R") > 0 Then
                                .DebutCycle = value(a)
                            End If
    
                            '--- contrôle de fin de cycle ---
                            If InStr(TabItemName(a), "M_Controle_FCY_R") > 0 Then
                                .ControleFinCycle = value(a)
                           End If
    
                            '--- numéro de la phase en cours ---
                            If InStr(TabItemName(a), "MW_Num_Phase_EC_R") > 0 Then
                                .NumPhaseEnCours = value(a)
                           End If
    
                            '--- temps de la phase en cours ---
                            If InStr(TabItemName(a), "MW_Tps_Phase_EC_R") > 0 Then
                                .TempsPhaseEnCours = value(a)
                           End If
    
                            '--- temps écoulé la phase en cours ---
                            If InStr(TabItemName(a), "MW_Tps_Ecoule_Ph_EC_R") > 0 Then
                                .TempsEcoulePhaseEnCours = value(a)
                            End If
    
                            '--- temps ajouté sur une intensité plus faible que prévue (panne de redresseur) ---
                            If InStr(TabItemName(a), "MW_Tps_Ajout_I_Faible_R") > 0 Then
                                .TempsAjouteSurIFaible = value(a)
                            End If
    
                            '--- temps totalisé ---
                            If InStr(TabItemName(a), "MW_Temps_Totalise_R") > 0 Then
                                .TempsTotalise = value(a)
                            End If
    
                            '------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- temps de la phase 1 ---
                            If InStr(TabItemName(a), "MW_Temps_Phase1_R") > 0 Then
                                .TDetailsPhases(1).TempsPhase = value(a)
                            End If
                            
                            '--- tension de la phase 1 ---
                            If InStr(TabItemName(a), "MW_U_Phase1_R") > 0 Then
                                .TDetailsPhases(1).UPhase = value(a)
                            End If
                            
                            '--- intensité de la phase 1 ---
                            If InStr(TabItemName(a), "MW_I_Phase1_R") > 0 Then
                                .TDetailsPhases(1).IPhase = value(a)
                            End If
    
                            '------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                            '--- temps de la phase 2 ---
                            If InStr(TabItemName(a), "MW_Temps_Phase2_R") > 0 Then
                                .TDetailsPhases(2).TempsPhase = value(a)
                            End If
    
                            '--- tension de la phase 2 ---
                            If InStr(TabItemName(a), "MW_U_Phase2_R") > 0 Then
                                .TDetailsPhases(2).UPhase = value(a)
                            End If
                            
                            '--- intensité de la phase 2 ---
                            If InStr(TabItemName(a), "MW_I_Phase2_R") > 0 Then
                                .TDetailsPhases(2).IPhase = value(a)
                            End If
    
                            '------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                            '--- temps de la phase 3 ---
                            If InStr(TabItemName(a), "MW_Temps_Phase3_R") > 0 Then
                                .TDetailsPhases(3).TempsPhase = value(a)
                            End If
                            
                            '--- tension de la phase 3 ---
                            If InStr(TabItemName(a), "MW_U_Phase3_R") > 0 Then
                                .TDetailsPhases(3).UPhase = value(a)
                            End If
                            
                            '--- intensité de la phase 3 ---
                            If InStr(TabItemName(a), "MW_I_Phase3_R") > 0 Then
                                .TDetailsPhases(3).IPhase = value(a)
                            End If
    
                            '------------------------------------------------------------------------------------------------------------------------------------------------------------
    
                            '--- temps de la phase 4 ---
                            If InStr(TabItemName(a), "MW_Temps_Phase4_R") > 0 Then
                                .TDetailsPhases(4).TempsPhase = value(a)
                            End If
                            
                            '--- tension de la phase 4 ---
                            If InStr(TabItemName(a), "MW_U_Phase4_R") > 0 Then
                                .TDetailsPhases(4).UPhase = value(a)
                            End If
                            
                            '--- intensité de la phase 4 ---
                            If InStr(TabItemName(a), "MW_I_Phase4_R") > 0 Then
                                .TDetailsPhases(4).IPhase = value(a)
                            End If
    
                            '------------------------------------------------------------------------------------------------------------------------------------------------------------
     
                            '--- numéro de charge traité par le redresseur ---
                            If InStr(TabItemName(a), "MW_Num_Charge_R") > 0 Then
                                .NumCharge = value(a)
                            End If
    
                            '--- temps théorique total de la gamme ---
                            If InStr(TabItemName(a), "MW_Tps_Total_Cycle_R") > 0 Then
                                .TempsTotalCycle = value(a)
                            End If
    
                            '--- temps restant ---
                            If InStr(TabItemName(a), "MW_Temps_Restant_R") > 0 Then
                                .TempsRestantCycle = value(a)
                            End If
    
                            '------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            '--- Ah ---
                            If InStr(TabItemName(a), "MD_Ah_R") > 0 Then
                                .Ah = value(a)
                            End If
                        
                        End With

                    End If

                End If
                
                '********************************************************************************************************
                '                                                        DEFAUTS REDRESSEURS
                '********************************************************************************************************
                If InStr(TabItemName(a), ".DEFAUTS_REDRESSEURS.") > 0 Then
                    
                    '--- analyse du dernier caractère pour trouver l'index du redresseur ---
                    Texte = Right(TabItemName(a), 1)
                    
                    If IsNumeric(Texte) = True Then
                    
                        '--- affectation de l'index ---
                        Index = CLng(Texte)
    
                        With TEtatsRedresseurs(Index)
    
                            '--- défaut général ---
                            If InStr(TabItemName(a), "M_Def_General_R") > 0 Then
                                .TEntreesAPI.M_DefautGeneral = IIf(value(a) = 0, False, True)
                            End If
                            
                            '--- délai trop long de mise en marche ---
                            If InStr(TabItemName(a), "M_DTL_Marche_R") > 0 Then
                                .TEntreesAPI.M_DelaiTropLongMarcheRedresseur = IIf(value(a) = 0, False, True)
                            End If
                            
                            '--- intensité non atteinte ---
                            If InStr(TabItemName(a), "M_Def_I_Non_Atteinte_R") > 0 Then
                                .TEntreesAPI.M_IntensiteNonAtteinte = IIf(value(a) = 0, False, True)
                            End If
                
                            '--- intensité instable ---
                            If InStr(TabItemName(a), "M_Def_I_Instable_R") > 0 Then
                                .TEntreesAPI.M_IntensiteInstable = IIf(value(a) = 0, False, True)
                            End If
                
                        End With
                
                    End If
                
                End If
                
                '********************************************************************************************************
                '                                                           TEMPERATURES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".TEMPERATURES.TempCuve") > 0 Then

                    '--- affectation de l'index ---
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            .Temperatures.TempActuelle = value(a) / 10
                        End With
                    End If
                    
                        
                End If

                '********************************************************************************************************
                '                                                           ETATS 1 DES CUVES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ETATS_CUVES.Etats1Cuve") > 0 Then
                    
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                       With TEtatsCuves(Index)

                        '--- affectation ---
                        .API_Etats_1 = CBin(Val("&h" & CHex(value(a))))

                        '--- entrées ---
                        .TEntreesAPI.E_CouverclesOuverts = Bit(.API_Etats_1, POS_BIT_COUVERCLES_OUVERTS)
                        .TEntreesAPI.E_CouverclesFermes = Bit(.API_Etats_1, POS_BIT_COUVERCLES_FERMES)
                        
                        '--- sorties ---
                        .TSortiesAPI.S_Chauffage = Bit(.API_Etats_1, POS_BIT_CHAUFFAGE)
                        .TSortiesAPI.S_Dem_Chauffage = Bit(.API_Etats_1, POS_BIT_DEM_CHAUFFAGE)
                        
                        .TSortiesAPI.S_Refroidissement = Bit(.API_Etats_1, POS_BIT_REFROIDISSEMENT)
                        .TSortiesAPI.S_Dem_Refroidissement = Bit(.API_Etats_1, POS_BIT_DEM_REFROIDISSEMENT)
                        
                        .TSortiesAPI.S_Pompe = Bit(.API_Etats_1, POS_BIT_POMPE)
                        .TSortiesAPI.S_Dem_Pompe = Bit(.API_Etats_1, POS_BIT_DEM_POMPE)
                        
                        .TSortiesAPI.S_EVEau = Bit(.API_Etats_1, POS_BIT_EV_EAU)
                        .TSortiesAPI.S_Dem_EVEau = Bit(.API_Etats_1, POS_BIT_DEM_EV_EAU)
                        
                        .TSortiesAPI.S_EVOuvertureCouvercles = Bit(.API_Etats_1, POS_BIT_OUVERTURE_COUVERCLES)
                        .TSortiesAPI.S_Dem_EVOuvertureCouvercles = Bit(.API_Etats_1, POS_BIT_DEM_OUVERTURE_COUVERCLES)
                        
                        .TSortiesAPI.S_EVFermetureCouvercles = Bit(.API_Etats_1, POS_BIT_FERMETURE_COUVERCLES)
                        .TSortiesAPI.S_Dem_EVFermetureCouvercles = Bit(.API_Etats_1, POS_BIT_DEM_FERMETURE_COUVERCLES)
                        
                        .TSortiesAPI.S_AgitationBain = Bit(.API_Etats_1, POS_BIT_AGITATION_BAIN)
                        .TSortiesAPI.S_Dem_AgitationBain = Bit(.API_Etats_1, POS_BIT_DEM_AGITATION_BAIN)

                        End With
                        '--- affectation des états de la cuve ---
                        AffectationEtatsCuve Index
                    End If
                    
                   
                
                    
                
                End If

                '********************************************************************************************************
                '                                                           ETATS 2 DES CUVES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ETATS_CUVES.Etats2Cuve") > 0 Then
                    
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)

                            '--- affectation ---
                            .API_Etats_2 = CBin(Val("&h" & CHex(value(a))))
    
                            '--- entrées ---
                            .TEntreesAPI.E_DefautChauffage = Bit(.API_Etats_2, POS_BIT_DEFAUT_CHAUFFAGE)
                            .TEntreesAPI.E_DefautRefroidissement = Bit(.API_Etats_2, POS_BIT_DEFAUT_REFROIDISSEMENT)
                            .TEntreesAPI.E_DefautPompe = Bit(.API_Etats_2, POS_BIT_DEFAUT_POMPE)
                            .TEntreesAPI.E_DefautEVEau = Bit(.API_Etats_2, POS_BIT_DEFAUT_EV_EAU)
                            .TEntreesAPI.E_DefautCouvercles = Bit(.API_Etats_2, POS_BIT_DEFAUT_COUVERCLES)
                            .TEntreesAPI.E_DefautAgitationBain = Bit(.API_Etats_2, POS_BIT_DEFAUT_AGITATION_BAIN)
                            
                            .TEntreesAPI.E_NiveauTresBas = Bit(.API_Etats_2, POS_BIT_NIVEAU_TRES_BAS)
                            .TEntreesAPI.E_NiveauIntermediaireBas = Bit(.API_Etats_2, POS_BIT_NIVEAU_INTERMEDIAIRE_BAS)
                            .TEntreesAPI.E_NiveauIntermediaireHaut = Bit(.API_Etats_2, POS_BIT_NIVEAU_INTERMEDIAIRE_HAUT)
                            .TEntreesAPI.E_NiveauTresHaut = Bit(.API_Etats_2, POS_BIT_NIVEAU_TRES_HAUT)
                            
                            .TEntreesAPI.E_ManuAutoRegulation = Bit(.API_Etats_2, POS_BIT_MANU_AUTO_REGULATION)
                            
                            '--- sorties ---
                            .TSortiesAPI.S_EVOuvertureCouvercles = Bit(.API_Etats_2, POS_BIT_OUVERTURE_COUVERCLES)
                            .TSortiesAPI.S_EVFermetureCouvercles = Bit(.API_Etats_2, POS_BIT_FERMETURE_COUVERCLES)
                            .TSortiesAPI.S_AgitationBain = Bit(.API_Etats_2, POS_BIT_AGITATION_BAIN)
    
                        End With
                        
                        '--- affectation des états de la cuve ---
                        AffectationEtatsCuve Index
                    End If
                    
                    
                
                End If

                '********************************************************************************************************
                '                                        MODES DE PRODUCTION DES CHAUFFAGES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".MODES_CHAUFFAGES_CUVES.ModeChauffagePCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur >= MODES_PRODUCTION.M_ARRET And Valeur <= MODES_PRODUCTION.M_PRODUCTION Then
                                .API_ModeProduction = Valeur
                                .API_Changements = True
                            End If
                        End With
                    
                    End If
                    
                    
                End If
        
                '********************************************************************************************************
                '                                                TEMPERATURES DE VEILLE
                '********************************************************************************************************
                If InStr(TabItemName(a), ".TEMP_VEILLE_CUVES.TempVeillePCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur <> 0 Then
                                .Temperatures.TempVeille = Valeur / 10
                                .API_Changements = True
                            End If
                        End With
                    
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                              TEMPERATURES DE PRODUCTION
                '********************************************************************************************************
                If InStr(TabItemName(a), ".TEMP_PRODUCTION_CUVES.TempProductionPCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur <> 0 Then
                                .Temperatures.TempProduction = Valeur / 10
                                .API_Changements = True
                            End If
                        End With
                    
                    End If
                    
                    
                End If
                
                '********************************************************************************************************
                '                                             ECARTS INFERIEUR DE REGULATION
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ECARTS_INF_REGUL_CUVES.EcartInfRegulPCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur <> 0 Then
                                .Temperatures.EcartInferieurRegul = Valeur / 10
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                             ECARTS SUPERIEUR DE REGULATION
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ECARTS_SUP_REGUL_CUVES.EcartSupRegulPCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur <> 0 Then
                                .Temperatures.EcartSuperieurRegul = Valeur / 10
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                                ECARTS INFERIEUR D'ALARME
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ECARTS_INF_ALARME_CUVES.EcartInfAlarmePCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur <> 0 Then
                                .Temperatures.EcartInferieurAlarme = Valeur / 10
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                                ECARTS SUPERIEUR D'ALARME
                '********************************************************************************************************
                If InStr(TabItemName(a), ".ECARTS_SUP_ALARME_CUVES.EcartSupAlarmePCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                         With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur <> 0 Then
                                .Temperatures.EcartSuperieurAlarme = Valeur / 10
                                .API_Changements = True
                            End If
                        End With
                    End If
                   
                End If
                
                '********************************************************************************************************
                '                                                       MODES DES POMPES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".MODES_POMPES_CUVES.ModePompePCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur >= MODES_POMPES.M_AUTO And Valeur <= MODES_POMPES.M_FORCER_MARCHE Then
                                .API_ModePompe = Valeur
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                                CYCLE EN AUTO DES POMPES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".CYCLES_AUTO_POMPES_CUVES.CycleAutoPompePCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur >= CYCLES_POMPES.CP_ARRET And Valeur <= CYCLES_POMPES.CP_MARCHE Then
                                .API_CyclePompe = Valeur
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                                     MODES DES COUVERCLES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".MODES_COUVERCLES_CUVES.ModeCouverclesPCCuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur >= MODES_COUVERCLES.M_AUTO And Valeur <= MODES_COUVERCLES.M_FORCER_OUVERTURE Then
                                .API_ModeCouvercles = Valeur
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
                '********************************************************************************************************
                '                                             CYCLE EN AUTO DES COUVERCLES
                '********************************************************************************************************
                If InStr(TabItemName(a), ".CYCLES_AUTO_COUVERCLES_CUVES.CycleAutoCouvercAPICuve") > 0 Then
                    Index = CLng(Right(TabItemName(a), 2))
                    Index = CORRESPONDANCES_IDX_CUVES_API(Index)
                    If (Index > 0) Then
                        With TEtatsCuves(Index)
                            Valeur = value(a)
                            If Valeur >= CYCLES_COUVERCLES.C_DEMANDE_FERMETURE And Valeur <= CYCLES_COUVERCLES.C_DEMANDE_OUVERTURE Then
                                .API_CycleCouvercles = Valeur
                                .API_Changements = True
                            End If
                        End With
                    End If
                    
                End If
                
        End Select

    Next a

    '********************************************************************************************************
    '                                             MODIFICATION DU CYCLE DU PONT 1
    '********************************************************************************************************
    If ModificationCycleP1 = True Then
        AnalyseCyclesPonts PONTS.P_1
        ModificationCycleP1 = False
    End If

    '********************************************************************************************************
    '                                             MODIFICATION DU CYCLE DU PONT 2
    '********************************************************************************************************
    If ModificationCycleP2 = True Then
        AnalyseCyclesPonts PONTS.P_2
        ModificationCycleP2 = False
    End If

    '********************************************************************************************************
    '                                              MODIFICATION SUR LES CHARIOTS
    '********************************************************************************************************
    If ModificationSurChariots = True Then

        If API_PresenceChariots <> String(16, 0) And API_VerrouillageChariots <> String(16, 0) Then

            For b = 0 To 3

                '--- chariots de chargement 1 à  4 ---
                If Bit(API_PresenceChariots, POS_BIT_C1 + b) = 1 Then
                    TEtatsPostes(POSTES.P_CHGT_1 + b).EtatsChariots = ETATS_CHARIOTS.E_PRESENT
                    If Bit(API_VerrouillageChariots, POS_BIT_C1 + b) = 1 Then
                        TEtatsPostes(POSTES.P_CHGT_1 + b).EtatsChariots = ETATS_CHARIOTS.E_PRESENT_VERROUILLE
                    End If
                Else
                    TEtatsPostes(POSTES.P_CHGT_1 + b).EtatsChariots = ETATS_CHARIOTS.E_ABSENT
                End If

                '--- chariots de déchargement 1 à 2 ---
                If b < 2 Then
                    If Bit(API_PresenceChariots, POS_BIT_D1 + b) = 1 Then
                        TEtatsPostes(POSTES.P_D1 + b).EtatsChariots = ETATS_CHARIOTS.E_PRESENT
                        If Bit(API_VerrouillageChariots, POS_BIT_D1 + b) = 1 Then
                            TEtatsPostes(POSTES.P_D1 + b).EtatsChariots = ETATS_CHARIOTS.E_PRESENT_VERROUILLE
                        End If
                    Else
                        TEtatsPostes(POSTES.P_D1 + b).EtatsChariots = ETATS_CHARIOTS.E_ABSENT
                    End If
                End If

            Next b

        End If

        '--- anti-rebond ---
        ModificationSurChariots = False

    End If

    '********************************************************************************************************
    '                                       ANALYSE SUR LES SENS DE MOUVEMENTS
    '********************************************************************************************************
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())

        With TEtatsPonts(a)

            '--- analyse des acrrcoches ---
            Select Case .ModePont

                Case MODES_PONTS.M_MAINTENANCE, _
                         MODES_PONTS.M_MANUEL, _
                         MODES_PONTS.M_SEMI_AUTOMATIQUE, _
                         MODES_PONTS.M_AUTOMATIQUE
                    '--- analyse des accroches ---
                    If .TSortiesAPI.S_EVMonteeAccroches = True And Not (.TEntreesAPI.M_AccrochesEnHaut) Then .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_VERS_HAUT
                    If .TSortiesAPI.S_EVDescenteAccroches = True And Not (.TEntreesAPI.M_AccrochesEnBas) Then .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_VERS_BAS
                    If .TEntreesAPI.M_AccrochesEnHaut = True Then .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_HAUT
                    If .TEntreesAPI.M_AccrochesEnBas = True Then .EtatsAccrochesCharge = ETATS_ACCROCHES.E_ACCROCHES_EN_BAS
                    
                Case Else
            End Select

        End With
    
    Next a

    '--- UNIQUEMENT POUR VISUALISATION ---
    'If TEtatsPonts(PONTS.P_1).TEntreesAPI.M_MoteurTourneTrlPont = True Then
    '    Debug.Print "Translation P1"
    'Else
    '    Debug.Print "Arrêt Translation P1"
    'End If
    'If TEtatsPonts(PONTS.P_1).TEntreesAPI.M_MoteurTourneLevPont = True Then
    '    Debug.Print "Levage P1"
    'Else
    '    Debug.Print "Arrêt Levage P1"
    'End If
    'If TEtatsPonts(PONTS.P_2).TEntreesAPI.M_MoteurTourneTrlPont = True Then
    '    Debug.Print "Translation P2"
    'Else
    '    Debug.Print "Arrêt Translation P2"
    'End If
    'If TEtatsPonts(PONTS.P_2).TEntreesAPI.M_MoteurTourneLevPont = True Then
    '    Debug.Print "Levage P2"
    'Else
    '    Debug.Print "Arrêt Levage P2"
    'End If

End Sub

Private Sub CBAcquittementAlarmes_Click()
    On Error Resume Next
    AcquittementAlarmes
End Sub

Private Sub LDonneesTransmisesAPI_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim OccForm As Form

    '--- recherche de la fenêtre ayant un affichage de données transmises vers API ---
    For Each OccForm In Forms
        Select Case OccForm.Name
            
            Case OccFChargementPrevisionnel.Name
                With OccFChargementPrevisionnel.LDonneesTransmisesAPI
                    .Caption = LDonneesTransmisesAPI.Caption                            'affichage
                    .Refresh                                                                                      'rafraichissement
                End With
            
            Case Else
        
        End Select
    Next

End Sub

Private Sub LMessages_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la fenêtre de la traçabilité des alarmes ---
    AppelFenetre FENETRES.F_TRACABILITE_ALARMES

End Sub

Private Sub MDIForm_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Static MemPassage As Boolean
    Dim a As Integer
    Dim EtatBusAPI As Integer                   'retourne état du bus de l'automate
    Dim ValeurRetourneeAPI As Long        'valeur retournée par une fonction concernant le dialogue avec l'automate
    
    '--- chargement de la configuration ---
    If MemPassage = False Then
        
        '--- affectation ---
        MemPassage = True       'pour éviter plusieurs passages même si la fenêtre devient
                                                 'active avant la fin de cette zone de démarrage

        '--- titre de la fenêtre ---
        Me.Caption = INDICATIF_PROGRAMME & " version : " & App.Major & "." & App.Minor & "." & App.Revision
        
        '--- affectation des images pour les barres d'outils ---
        For a = TOBOutils.LBound To TOBOutils.UBound
            Set TOBOutils(a).ImageList = ILOutils
            Set TOBOutils(a).HotImageList = ILOutils
        Next a
        
        '--- ouverture de la fenetre d'analyse ---
        FAnalyseDeDemarrage.Show vbModeless
        FAnalyseDeDemarrage.Refresh
        DoEvents

        '--- initialisation des variables ---
        InitialisationVariables
            
        '--- configuration ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement de la configuration")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeConfiguration())
        
        '--- affectation des chemins ---
        Select Case TypePC
            Case TYPES_PC.PC_SUR_LIGNE
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Affectation des chemins (PC de l'atelier d'anodisation)")
            Case TYPES_PC.PC_ENTREPRISE
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Affectation des chemins (Un des PC de l'entreprise)")
            Case TYPES_PC.PC_DISTANT
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Affectation des chemins (PC Distant)")
            Case Else
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "TYPE de PC INCONNU (Voir fichier de configuration")
        End Select
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, AffectationChemins)
        
        '--- UNIQUEMENT POUR LA CONSTRUCTION DES FICHIERS ---
        'Bidon = SauveProgCyclique()
        'Bidon = SauveRegulation()
        'Bidon = SauveJourneesTypes()


        'Call testRecord
        
        '--- initialisation des images ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Initialisation des images")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, InitialisationImages())
        
        '--- chargement des défauts ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des défauts")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeDefauts())
        
        '--- attribution des n° de défauts ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Attribution des numéros de défauts")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, AttributionNumDefauts())
        
        '--- chargement des zones ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des zones de la ligne")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeZones())
        
         '--- chargement des zones ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des barres de la ligne")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeBarres())
        
        '--- chargement des caractéristiques des cuves ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des cuves")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeCuves())
        
        '--- chargement des types de gammes ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des matières")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeMatieres())
        
        '--- chargement des caractéristiques des postes ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des postes")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargePostes())


        '--- chargement des paramètres ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des paramètres")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeParametres())
       

        '--- chargement des codes des actions ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des actions")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeActions())
        
        '--- chargement des temps de mouvements ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des temps de mouvements")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeTempsMouvements())
        
        '--- chargement des prémisses ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des prémisses")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargePremisses())
        
        '--- chargement de la régulation ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement du fichier de la régulation")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeRegulation())
        
        '--- chargement des journées types ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement du fichier des journées types")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeJourneesTypes())
    
        '--- chargement du programmateur cyclique ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement du fichier du programmateur cyclique")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeProgCyclique())
    
        '--- chargement des annexes ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement du fichier des annexes (ventilation, etc ...)")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeAnnexes())
    
        '--- chargement des caractéristiques des redresseurs ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des redresseurs")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeRedresseurs())
        
        '--- chargement des paramètres de la ligne ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Chargement des paramètres de la ligne")
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, ChargeParametresLigne())
    
        If PROGRAMME_AVEC_AUTOMATE = True Then
            
            '--- ouverture de la liaison automate ---
            Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Initialisation du réseau automate")
            Call initbus(EtatBusAPI)
            If EtatBusAPI = 0 Then
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "")
            Else
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "INCIDENT")
            End If
            
            '--- activation de la configuration principale avec l'automate ---
            Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Activation de la configuration API principale")
            ValeurRetourneeAPI = ActiveConfiguration(Me.AOCFPrincipale)
            If ValeurRetourneeAPI = 0 Then
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "")
            Else
                Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, MessagesEtatCommunication(ValeurRetourneeAPI))
            End If
                    
            '--- référencement du serveur de communication APPLICOM ---
            Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Référencement du serveur de communication APPLICOM")
            RefServeur = OccFPrincipale.AOCFPrincipale.GetServerRef(NOM_SERVEUR_APPLICOM)
            If RefServeur > 0 Then
               Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "")
            Else
               Select Case RefServeur
                   Case -1: Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "Le nom du serveur passé en paramètre n'existe pas")
                   Case -7: Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "Le paramètre est invalide, il ne contient pas une chaîne de caractères")
                   Case Else: Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "ERREUR D'ACCES AU SERVEUR APPLICOM")
              End Select
            End If
        
        End If
        
        '--- affectation des canaux pour les fichiers ---
        For a = REDRESSEURS.R_C13 To REDRESSEURS.R_C16
            TCanauxFichiersTraçabilite(a) = CANAL_DEPART_TRACABILITE + a
        Next a
        
        '--- premier affichage de la date et de l'heure, lancement du timer ---
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Affichage de la date et de l'heure")
        AfficheDateHeure
        TimerDateHeure.Enabled = True
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "")
        
        '--- lancement des timers ---
        TimerLigneAlarmes.Enabled = True

        '--- lancement du noyau multi-tâches ---
        TimerNoyauCentral.Enabled = True
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_LIBELLE, "Lancement du noyau multi-tâches")
        While PremierPassageNoyauCentral = False
            DoEvents
        Wend
        Call FAnalyseDeDemarrage.ControleFonction(AFFICHAGE_ANALYSE, "")
        
        '--- fermeture de la fenetre d'analyse ---
        If PROGRAMME_TERMINE = True Then Call Sleep(1500)
        Unload FAnalyseDeDemarrage
        Set FAnalyseDeDemarrage = Nothing
        
        '--- lancement du synoptique ---
        AppelFenetre FENETRES.F_SYNOPTIQUE
        
        

        
    End If

End Sub

Private Sub MDIForm_Load()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    
    '--- contrôle d'une autre instance du programme ---
    If App.PrevInstance = True Then
        
        Bidon = AppelFenetre(F_MESSAGE, _
                                           "Programme Anodisation", _
                                          vbCrLf & vbCrLf & "c|Le programme" & vbCrLf & _
                                          "mc|" & App.EXEName & ".EXE" & vbCrLf & _
                                          "cs|est déjà chargé", _
                                           TYPES_MESSAGES.T_ATTENTION, _
                                           TYPES_BOUTONS.T_CONFIRMER, _
                                            EMPLACEMENT_FOCUS.E_SUR_CONFIRMER)
        End
    
    Else
    
        '--- chargement des dernières positions de la fenetre principale ---
        Me.Left = GetSetting(App.Title, Me.Name, "Distance bord gauche", 1000)
        Me.Top = GetSetting(App.Title, Me.Name, "Distance bord haut", 1000)
        Me.Width = GetSetting(App.Title, Me.Name, "Largeur", 6500)
        Me.Height = GetSetting(App.Title, Me.Name, "Hauteur", 6500)
    
    End If
    

    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim OccForm As Form
    
    '--- demande de confirmation ---
    'If MotDePasse(True) = True Then

        '--- demande de confirmation ---
        If AppelFenetre(F_MESSAGE, _
                                TITRE_MESSAGES, _
                                "Les gammes et la traçabilité seront suspendues." & vbCrLf & _
                                "Le programme ne fonctionnera plus car la sortie est" & vbCrLf & _
                                "définitive." & _
                                vbCrLf & vbCrLf & vbCrLf & _
                                "cs|Voulez-vous réellement quitter le programme ?", _
                                TYPES_MESSAGES.T_ATTENTION, _
                                TYPES_BOUTONS.T_OUI_NON, _
                                EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
            
            '--- fermeture de toutes les fenetres ---
            For Each OccForm In Forms
                Select Case OccForm.Name
                    Case OccFPrincipale.Name
                    Case Else
                        Unload OccForm
                        Set OccForm = Nothing
                End Select
            Next
            
            '--- sauvegarde de la configuration ---
            'SauveConfiguration
            
        Else
    
            '--- affectation ---
            Cancel = True
    
        End If
    
    'End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next

    '--- sauvegarde des dernières positions de la fenêtre principale ---
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, Me.Name, "Distance bord gauche", Me.Left
        SaveSetting App.Title, Me.Name, "Distance bord haut", Me.Top
        SaveSetting App.Title, Me.Name, "Largeur", Me.Width
        SaveSetting App.Title, Me.Name, "Hauteur", Me.Height
    End If
    
    
    
    
    modMultiThreading.Uninitialize
    MBaseDeDonnees.mcInsertClipper = Nothing
    TerminateThread MBaseDeDonnees.mlID, 0
End Sub

Private Sub Coller()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- analyse en fonction du contrôle ---
    With Clipboard
        If TypeOf Screen.ActiveControl Is TextBox Then
            Screen.ActiveControl.SelText = .GetText
        ElseIf TypeOf Screen.ActiveControl Is RichTextBox Then
            Screen.ActiveControl.SelRTF = .GetText
        End If
    End With
    
End Sub

Private Sub Copier()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- analyse en fonction du contrôle ---
    With Clipboard
        If TypeOf Screen.ActiveControl Is TextBox Then
            .Clear
            .SetText Screen.ActiveControl.SelText
        ElseIf TypeOf Screen.ActiveControl Is RichTextBox Then
            .Clear
            .SetText Screen.ActiveControl.SelRTF
        End If
    End With

End Sub

Private Sub Couper()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- analyse en fonction du contrôle ---
    With Clipboard
        If TypeOf Screen.ActiveControl Is TextBox Then
            .Clear
            .SetText Screen.ActiveControl.SelText
            Screen.ActiveControl.SelText = ""
        ElseIf TypeOf Screen.ActiveControl Is RichTextBox Then
            .Clear
            .SetText Screen.ActiveControl.SelRTF
            Screen.ActiveControl.SelRTF = ""
        End If
    End With
    
End Sub

Private Sub MenuAPropos_Click()
    On Error Resume Next
    AppelFenetre F_APROPOS
End Sub

Private Sub MenuDiversChargementCheminBDCLIPPER_Click()
    On Error Resume Next
    Bidon = ChargeCheminBDCLIPPER()
End Sub

Private Sub MenuDiversNettoyageGraphesProduction_Click()
    On Error Resume Next
    AppelFenetre F_NETTOYAGE_GRAPHES_PRODUCTION
End Sub

Private Sub MenuFenetresActualiser_Click()
    On Error Resume Next
    OccFPrincipale.ActiveForm.Refresh
    Call OccFSynoptique.GestionImageTampon(True)
End Sub

Private Sub MenuFenetresEnCascade_Click()
    On Error Resume Next
    OccFSynoptique.Move 0, 0
    Me.Arrange vbCascade
End Sub

Private Sub MenuFenetresFermerTout_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim OccForm As Form

    For Each OccForm In Forms
        Select Case OccForm.Name
            Case OccFPrincipale.Name, OccFSynoptique.Name
            Case Else
                Unload OccForm
                Set OccForm = Nothing
        End Select
    Next

    '--- déplacement du synoptique si nécessaire ---
    OccFSynoptique.Move 0, 0
    Me.Arrange vbCascade

End Sub

Private Sub MenuFenetresMosaiqueCalculee_Click()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim Cptfenetres As Integer, _
            CptSurAxeY As Integer
    Dim OccForm As Form
    
    '--- affectation ---
    
    '--- déplacement du synoptique ---
    OccFSynoptique.Move 0, 0

    For Each OccForm In Forms
        Select Case OccForm.Name
            Case OccFPrincipale.Name, OccFSynoptique.Name
            Case Else
                With OccForm
                    
                    '--- dimensionnement ---
                    .WindowState = vbNormal
                    .Height = OccFSynoptique.Height
                    .Width = OccFSynoptique.Width
                    .Refresh

                    '--- affectation ---
                    Inc Cptfenetres
                    .Top = CptSurAxeY * OccFSynoptique.Height
                    Select Case Cptfenetres
                        Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35, 37
                            .Left = OccFSynoptique.Width
                            Inc CptSurAxeY
                        Case Else
                            .Left = 0
                    End Select
        
                End With
        End Select
    Next

End Sub

Private Sub MenuFenetresMosaiqueHorizontale_Click()
    On Error Resume Next
    OccFSynoptique.Move 0, 0
    Me.Arrange vbTileHorizontal
End Sub

Private Sub MenuFenetresMosaiqueVerticale_Click()
    On Error Resume Next
    OccFSynoptique.Move 0, 0
    Me.Arrange vbTileVertical
End Sub

Private Sub MenuQuitter_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub TimerDateHeure_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine ---
    TimerDateHeure.Enabled = False
    AfficheDateHeure
    TimerDateHeure.Enabled = True

End Sub

Private Sub TimerLigneAlarmes_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerLigneAlarmes.Enabled = False
    VisualisationLigneAlarmes
    TimerLigneAlarmes.Enabled = True
    
End Sub

Private Sub TimerNoyauCentral_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Static CptCycle As Integer
    
    '--- appel du noyau central ---
    TimerNoyauCentral.Enabled = False
    If CptCycle <= 60 Then
        NoyauCentral
        Inc CptCycle
    Else
        With OccFPrincipale.LTempsNoyauCentral
            .Caption = "S"
            .BackColor = COULEURS.BLEU_3
            .ForeColor = COULEURS.JAUNE_3
            .Refresh
        End With
        DoEvents
        CptCycle = 0
    End If
    TimerNoyauCentral.Enabled = True
    
End Sub

Private Sub TOBOutils_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sélection en fonction de l'outil cliqué ---
    Select Case Button.Key
        
        Case "AperçuAvantImpression"
            '--- appel de la routine gérant les aperçus avant impression ---
            Impressions TYPES_IMPRESSIONS.TI_APERCU_AVANT_IMPRESSION

        Case "Calculatrice"
            '--- calculatrice ---
            AppelCalculatrice

        Case "OrganisationLigne"
            '--- organisation de la ligne ---
            AppelFenetre FENETRES.F_ORGANISATION_LIGNE
            
        Case "MoteurInference"
            '--- moteur d'inférence ---
            AppelFenetre FENETRES.F_MOTEUR_INFERENCE
            
        Case "ModeCyclique"
            '--- mode cyclique ---
            AppelFenetre FENETRES.F_MODE_CYCLIQUE
                
        Case "GammesAnodisation"
            '--- gammes de production ---
            AppelFenetre FENETRES.F_GAMMES_ANODISATION
            
        Case "Tracabilite"
            '--- traçabilité de production ---
            AppelFenetre FENETRES.F_TRACABILITE_PRODUCTION
            
        Case "ChargesEnLigne"
            '--- charges en ligne ---
            AppelFenetre FENETRES.F_CHARGES_EN_LIGNE
            
        Case "CyclesPonts"
            '--- cycles des ponts ---
            AppelFenetre FENETRES.F_CYCLES_PONTS
            
        Case "ChargementPrevisionnel"
            '--- chargement / prévisionnel ---
            AppelFenetre FENETRES.F_CHARGEMENT_PREVISIONNEL
            
        Case "Redresseurs"
            '--- redresseurs ---
            AppelFenetre FENETRES.F_GESTION_REDRESSEURS
            
        Case "Cuves"
            '--- gestion des cuves ---
            AppelFenetre FENETRES.F_GESTION_CUVES
            
        Case "Regulation"
            '--- gestion de la régulation ---
            AppelFenetre FENETRES.F_GESTION_REGULATION
            
        Case "ProgrammateurCyclique"
            '--- programmateur cyclique ---
            AppelFenetre FENETRES.F_PROGRAMMATEUR_CYCLIQUE
            
        Case "Annexes"
            '--- annexes ---
            AppelFenetre FENETRES.F_ANNEXES
            
        Case "Defauts"
            '--- liste des défauts ---
            AppelFenetre FENETRES.F_LISTE_DEFAUTS
        
        Case "Maintenance"
            '--- maintenance ---
            AppelFenetre FENETRES.F_MAINTENANCE

        Case "Fermer tout"
            '--- fermeture de toutes les fenêtres ---
            Call MenuFenetresFermerTout_Click
        
        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affectation des états d'une cuve
' Entrées : NumCuve -> Numéro d'une cuve
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffectationEtatsCuve(ByVal NumCuve As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim EtatsCouvercles As ETATS_COUVERCLES                     'états des couvercles
    
    If NumCuve >= LBound(TEtatsCuves()) And NumCuve <= UBound(TEtatsCuves()) Then

        With TEtatsCuves(NumCuve)

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
            '--- mode de la régulation ---
            If .TEntreesAPI.E_ManuAutoRegulation = False Then
                .ModeRegulation = MODES_REGULATION.MR_MANUEL
            Else
                .ModeRegulation = MODES_REGULATION.MR_AUTOMATIQUE
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- états du chauffage ---
            If .DefinitionCuve.NbrChauffages > 0 Then
                If .TEntreesAPI.E_DefautChauffage = True Then
                    .EtatsChauffage = ETATS_CHAUFFAGES.M_DEFAUT
                Else
                    If .TSortiesAPI.S_Chauffage = False Then
                        .EtatsChauffage = ETATS_CHAUFFAGES.M_ARRET
                    Else
                        .EtatsChauffage = ETATS_CHAUFFAGES.M_MARCHE
                    End If
                End If
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- états du refroidissement d'un bain ---
            If .DefinitionCuve.PresenceRefroidissementBain = True Then
                If .TEntreesAPI.E_DefautRefroidissement = True Then
                    .EtatsRefroidissementBain = ETATS_REFROIDISSEMENT_BAIN.M_DEFAUT
                Else
                    If .TSortiesAPI.S_Refroidissement = False Then
                        .EtatsRefroidissementBain = ETATS_REFROIDISSEMENT_BAIN.M_ARRET
                    Else
                        .EtatsRefroidissementBain = ETATS_REFROIDISSEMENT_BAIN.M_MARCHE
                    End If
                End If
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- états de la pompe ---
            If .DefinitionCuve.PresencePompe = True Then
                If .TEntreesAPI.E_DefautPompe = True Then
                    .EtatsPompe = ETATS_POMPES.E_DEFAUT
                Else
                    If .TSortiesAPI.S_Pompe = False Then
                        .EtatsPompe = ETATS_POMPES.E_ARRET
                    Else
                        .EtatsPompe = ETATS_POMPES.E_MARCHE
                    End If
                End If
            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- niveaux ---
            If .TEntreesAPI.E_NiveauTresBas = False And _
                .TEntreesAPI.E_NiveauIntermediaireBas = False And _
                .TEntreesAPI.E_NiveauIntermediaireHaut = False And _
                .TEntreesAPI.E_NiveauTresHaut = False Then
            
                .EtatsNiveaux = ETATS_NIVEAUX.E_NORMAL
            
            Else
            
                If .TEntreesAPI.E_NiveauIntermediaireBas = True Then
                    .EtatsNiveaux = ETATS_NIVEAUX.E_INTERMEDIAIRE_BAS
                End If
                If .TEntreesAPI.E_NiveauIntermediaireHaut = True Then
                    .EtatsNiveaux = ETATS_NIVEAUX.E_INTERMEDIAIRE_HAUT
                End If
                If .DefinitionCuve.PresenceNiveauHaut = True Then
                    If .TEntreesAPI.E_NiveauTresHaut = True Then
                        .EtatsNiveaux = ETATS_NIVEAUX.E_TRES_HAUT
                    End If
                End If
                If .DefinitionCuve.PresenceNiveauBas = True Then
                    If .TEntreesAPI.E_NiveauTresBas = True Then
                        .EtatsNiveaux = ETATS_NIVEAUX.E_TRES_BAS
                    End If
                End If

            End If

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

            '--- électro-vanne d'arrivée d'eau ---
            If .DefinitionCuve.PresenceEVEau = True Then
                If .TEntreesAPI.E_DefautEVEau = True Then
                    .EtatsEVEau = ETATS_EV_EAU.E_DEFAUT
                ElseIf .TEntreesAPI.E_DelaiTropLongEVEau = True Then
                    .EtatsEVEau = ETATS_EV_EAU.E_DELAI_LONG
                Else
                    If .TSortiesAPI.S_EVEau = True Then
                        .EtatsEVEau = ETATS_EV_EAU.E_OUVERTE
                    Else
                        .EtatsEVEau = ETATS_EV_EAU.E_FERMEE
                    End If
                End If
            End If

            '--- analyse pour les couvercles et transfert des valeurs dans les postes concernés ---
            'If .TSortiesAPI.S_EVOuvertureCouvercles = True And Not (.TEntreesAPI.E_CouverclesOuverts) Then EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_EN_OUVERTURE
            'If .TSortiesAPI.S_EVFermetureCouvercles = True And Not (.TEntreesAPI.E_CouverclesFermes) Then EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_EN_FERMETURE
            'If .TEntreesAPI.E_CouverclesOuverts = True Then EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_OUVERTS
            'If .TEntreesAPI.E_CouverclesFermes = True Then EtatsCouvercles = ETATS_COUVERCLES.E_COUVERCLES_FERMES
            'If .TEntreesAPI.E_DefautCouvercles = True Then EtatsCouvercles = ETATS_COUVERCLES.E_DEFAUT_COUVERCLES
            'Select Case Index
            '    Case CUVES_API.C_A1: TEtatsPostes(POSTES.P_A1).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_A2: TEtatsPostes(POSTES.P_A2).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_A13: TEtatsPostes(POSTES.P_A13).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_A14: TEtatsPostes(POSTES.P_A14).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_A17: TEtatsPostes(POSTES.P_A17).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_A18: TEtatsPostes(POSTES.P_A18).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_B1: TEtatsPostes(POSTES.P_B1).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_B4: TEtatsPostes(POSTES.P_B4).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_B5: TEtatsPostes(POSTES.P_B5).EtatsCouvercles = EtatsCouvercles
            '    Case CUVES_API.C_B11: TEtatsPostes(POSTES.P_B11).EtatsCouvercles = EtatsCouvercles
            '    Case Else
            'End Select

        End With

    End If
                
End Sub

Private Sub TOBOutils_ButtonMenuClick(Index As Integer, ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sélection en fonction de l'outil cliqué ---
    Select Case ButtonMenu.Key
        
        Case "ImprimerDirectement"
            '--- appel de la routine gérant l'impression ---
            Impressions TYPES_IMPRESSIONS.TI_IMPRIMER
                
        Case "ImprimerFenetreActive"
            '--- appel de la routine gérant les impressions ---
            Impressions TYPES_IMPRESSIONS.TI_IMPRIMER_FENETRE_ACTIVE

        Case "Premisses"
               '--- prémisses ---
            AppelFenetre F_PREMISSES
            
        Case "TempsMouvements"
            '--- temps de mouvements ---
            AppelFenetre F_TEMPS_MOUVEMENTS
        
        Case Else
    End Select

End Sub

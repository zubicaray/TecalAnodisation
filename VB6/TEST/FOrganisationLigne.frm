VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FOrganisationLigne 
   ClientHeight    =   13290
   ClientLeft      =   1320
   ClientTop       =   2445
   ClientWidth     =   17145
   Icon            =   "FOrganisationLigne.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   13290
   ScaleWidth      =   17145
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      Picture         =   "FOrganisationLigne.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   17085
      TabIndex        =   14
      Top             =   0
      Width           =   17145
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ORGANISATION DE LA LIGNE"
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
         Height          =   225
         Left            =   240
         TabIndex        =   15
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
      ScaleWidth      =   17085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   12195
      Width           =   17145
      Begin VB.PictureBox PBBoutonsEnregistrements 
         Height          =   840
         Left            =   10080
         ScaleHeight     =   780
         ScaleWidth      =   5055
         TabIndex        =   9
         Top             =   120
         Width           =   5115
         Begin VB.CommandButton CBEnregistrements 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Premier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Index           =   0
            Left            =   0
            MaskColor       =   &H00FF00FF&
            Picture         =   "FOrganisationLigne.frx":24A8C
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   " Accès au premier enregistrement "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton CBEnregistrements 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Précédent"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Index           =   1
            Left            =   1260
            MaskColor       =   &H00FF00FF&
            Picture         =   "FOrganisationLigne.frx":24D76
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   " Enregistrement précédent "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton CBEnregistrements 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Suivant"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Index           =   2
            Left            =   2520
            MaskColor       =   &H00FF00FF&
            Picture         =   "FOrganisationLigne.frx":25060
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   " Enregistrement suivant "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton CBEnregistrements 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Dernier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   780
            Index           =   3
            Left            =   3780
            MaskColor       =   &H00FF00FF&
            Picture         =   "FOrganisationLigne.frx":2534A
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   " Accès au dernier enregistrement "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   4
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   5
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FOrganisationLigne.frx":25634
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
            Picture         =   "FOrganisationLigne.frx":257DE
            Style           =   1  'Graphical
            TabIndex        =   6
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
         DownPicture     =   "FOrganisationLigne.frx":25988
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
         Left            =   15480
         MaskColor       =   &H00FF00FF&
         Picture         =   "FOrganisationLigne.frx":2608A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1395
      End
      Begin VB.CommandButton CBActualiser 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Actualise&r"
         DownPicture     =   "FOrganisationLigne.frx":2678C
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
         Picture         =   "FOrganisationLigne.frx":26E8E
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   " Actualiser les données "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1395
      End
      Begin MSComctlLib.ImageList ILOutils 
         Left            =   2040
         Top             =   120
         _ExtentX        =   794
         _ExtentY        =   794
         BackColor       =   -2147483643
         ImageWidth      =   13
         ImageHeight     =   10
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrganisationLigne.frx":27590
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FOrganisationLigne.frx":27778
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   1440
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   14175
      Index           =   0
      Left            =   0
      ScaleHeight     =   14175
      ScaleWidth      =   17145
      TabIndex        =   7
      Top             =   375
      Width           =   17145
      Begin VB.PictureBox PBDeplacementFenetre 
         BorderStyle     =   0  'None
         Height          =   14055
         Index           =   1
         Left            =   0
         ScaleHeight     =   14055
         ScaleWidth      =   28725
         TabIndex        =   8
         Top             =   0
         Width           =   28725
         Begin VB.PictureBox PBMoteur 
            Height          =   5895
            Left            =   180
            Picture         =   "FOrganisationLigne.frx":2797C
            ScaleHeight     =   5835
            ScaleWidth      =   11655
            TabIndex        =   28
            Top             =   6660
            Width           =   11715
            Begin VB.CommandButton CBTransfertProchainCycleEnActuel 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Transfert du prochain cycle en cycle actuel - PONT 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   8820
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   2820
               Width           =   2535
            End
            Begin VB.CommandButton CBCycleActuelPont 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Cycle actuel - PONT 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   7320
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   480
               Width           =   2535
            End
            Begin VB.CommandButton CBPrevisionnel 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Prévisionnel"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   8220
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   5160
               Width           =   2535
            End
            Begin VB.CommandButton CBPremisses 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Les PREMISSES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1860
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   3480
               Width           =   3675
            End
            Begin VB.CommandButton CBProchainCyclePont 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Prochain cycle - PONT 2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   2
               Left            =   4560
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   5160
               Width           =   2535
            End
            Begin VB.CommandButton CBProchainCyclePont 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Prochain cycle - PONT 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   900
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   5160
               Width           =   2535
            End
            Begin VB.CommandButton CBMoteurInference 
               BackColor       =   &H00C0FFC0&
               Caption         =   "MOTEUR D'INFERENCE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   4200
               Width           =   3675
            End
            Begin VB.CommandButton CBAnalyseChargesEnCours 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Analyse des charges en ligne"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   3960
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   2820
               Width           =   3735
            End
            Begin VB.CommandButton CBTransfertProchainCycleEnActuel 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Transfert du prochain cycle en cycle actuel - PONT 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   300
               Style           =   1  'Graphical
               TabIndex        =   31
               Top             =   2820
               Width           =   2535
            End
            Begin VB.CommandButton CBCycleActuelPont 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Cycle actuel - PONT 1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Index           =   1
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   30
               Top             =   480
               Width           =   2535
            End
            Begin VB.CommandButton CBChargement 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Chargement"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   4560
               Style           =   1  'Graphical
               TabIndex        =   29
               Top             =   1140
               Width           =   2535
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   9
               Left            =   3000
               TabIndex        =   56
               Top             =   240
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2CF84
               maskpicture     =   "FOrganisationLigne.frx":2D196
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   10
               Left            =   8520
               TabIndex        =   57
               Top             =   240
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2D3A8
               maskpicture     =   "FOrganisationLigne.frx":2D5BA
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   11
               Left            =   1500
               TabIndex        =   58
               Top             =   2580
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2D7CC
               maskpicture     =   "FOrganisationLigne.frx":2D9DE
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   12
               Left            =   4500
               TabIndex        =   59
               Top             =   2580
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2DBF0
               maskpicture     =   "FOrganisationLigne.frx":2DE02
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   13
               Left            =   5760
               TabIndex        =   60
               Top             =   2580
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2E014
               maskpicture     =   "FOrganisationLigne.frx":2E226
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   14
               Left            =   7020
               TabIndex        =   61
               Top             =   2580
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2E438
               maskpicture     =   "FOrganisationLigne.frx":2E64A
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   15
               Left            =   10020
               TabIndex        =   62
               Top             =   2580
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2E85C
               maskpicture     =   "FOrganisationLigne.frx":2EA6E
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   16
               Left            =   5760
               TabIndex        =   63
               Top             =   3960
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2EC80
               maskpicture     =   "FOrganisationLigne.frx":2EE92
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   17
               Left            =   2100
               TabIndex        =   64
               Top             =   4920
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2F0A4
               maskpicture     =   "FOrganisationLigne.frx":2F2B6
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   18
               Left            =   5760
               TabIndex        =   65
               Top             =   4920
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2F4C8
               maskpicture     =   "FOrganisationLigne.frx":2F6DA
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   19
               Left            =   9420
               TabIndex        =   66
               Top             =   4920
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":2F8EC
               maskpicture     =   "FOrganisationLigne.frx":2FAFE
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFinCycleActuelPont 
               Height          =   1575
               Index           =   1
               Left            =   2280
               TabIndex        =   67
               Top             =   1200
               Width           =   1575
               _extentx        =   2778
               _extenty        =   2778
               picture         =   "FOrganisationLigne.frx":2FD10
               downpicture     =   "FOrganisationLigne.frx":37EFE
               maskpicture     =   "FOrganisationLigne.frx":400EC
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   20
               Left            =   3000
               TabIndex        =   68
               Top             =   960
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":482DA
               maskpicture     =   "FOrganisationLigne.frx":484EC
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFinCycleActuelPont 
               Height          =   1575
               Index           =   2
               Left            =   7800
               TabIndex        =   69
               Top             =   1200
               Width           =   1575
               _extentx        =   2778
               _extenty        =   2778
               picture         =   "FOrganisationLigne.frx":486FE
               downpicture     =   "FOrganisationLigne.frx":508EC
               maskpicture     =   "FOrganisationLigne.frx":58ADA
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   21
               Left            =   8520
               TabIndex        =   70
               Top             =   960
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":60CC8
               maskpicture     =   "FOrganisationLigne.frx":60EDA
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   22
               Left            =   4500
               TabIndex        =   71
               Top             =   3960
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":610EC
               maskpicture     =   "FOrganisationLigne.frx":612FE
               maskcolor       =   16711935
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "NON"
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
               Height          =   195
               Index           =   5
               Left            =   7320
               TabIndex        =   43
               Top             =   1740
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "OUI"
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
               Height          =   195
               Index           =   4
               Left            =   9420
               TabIndex        =   42
               Top             =   1740
               Width           =   375
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "NON"
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
               Height          =   195
               Index           =   3
               Left            =   3900
               TabIndex        =   41
               Top             =   1740
               Width           =   435
            End
            Begin VB.Label LLibelles 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "OUI"
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
               Height          =   195
               Index           =   2
               Left            =   1860
               TabIndex        =   40
               Top             =   1740
               Width           =   375
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   31
               X1              =   9480
               X2              =   9480
               Y1              =   4980
               Y2              =   4920
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   20
               X1              =   2160
               X2              =   2160
               Y1              =   4920
               Y2              =   5040
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   30
               X1              =   4560
               X2              =   4560
               Y1              =   4920
               Y2              =   4680
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   29
               X1              =   4560
               X2              =   2160
               Y1              =   4920
               Y2              =   4920
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   25
               X1              =   10080
               X2              =   10080
               Y1              =   3660
               Y2              =   3300
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   19
               X1              =   10080
               X2              =   10080
               Y1              =   2580
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   18
               X1              =   10080
               X2              =   9360
               Y1              =   1980
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   28
               X1              =   5820
               X2              =   5820
               Y1              =   4020
               Y2              =   3300
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   27
               X1              =   5820
               X2              =   5820
               Y1              =   4920
               Y2              =   4620
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   24
               X1              =   7020
               X2              =   7020
               Y1              =   4920
               Y2              =   4680
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   23
               X1              =   9480
               X2              =   7020
               Y1              =   4920
               Y2              =   4920
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   22
               X1              =   11520
               X2              =   11520
               Y1              =   3660
               Y2              =   240
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   21
               X1              =   8580
               X2              =   11520
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   17
               X1              =   120
               X2              =   1560
               Y1              =   3660
               Y2              =   3660
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   14
               X1              =   7080
               X2              =   7080
               Y1              =   2580
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   13
               X1              =   4560
               X2              =   4560
               Y1              =   2580
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   12
               X1              =   5820
               X2              =   5820
               Y1              =   2640
               Y2              =   1620
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   10
               X1              =   120
               X2              =   120
               Y1              =   3660
               Y2              =   240
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   8
               X1              =   1560
               X2              =   1560
               Y1              =   2640
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   7
               X1              =   1560
               X2              =   1560
               Y1              =   3300
               Y2              =   3660
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   1
               X1              =   4560
               X2              =   3840
               Y1              =   1980
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   36
               X1              =   2280
               X2              =   1560
               Y1              =   1980
               Y2              =   1980
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   37
               X1              =   10080
               X2              =   11520
               Y1              =   3660
               Y2              =   3660
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   11
               X1              =   3060
               X2              =   120
               Y1              =   240
               Y2              =   240
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H000000FF&
               BorderWidth     =   3
               Index           =   16
               X1              =   7800
               X2              =   7080
               Y1              =   1980
               Y2              =   1980
            End
         End
         Begin VB.PictureBox PBCaracteristiquesLigne 
            Height          =   5715
            Left            =   180
            Picture         =   "FOrganisationLigne.frx":61510
            ScaleHeight     =   5655
            ScaleWidth      =   11655
            TabIndex        =   18
            Top             =   480
            Width           =   11715
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   0
               Left            =   1260
               TabIndex        =   47
               Top             =   4620
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":65E37
               maskpicture     =   "FOrganisationLigne.frx":66049
               maskcolor       =   16711935
            End
            Begin VB.CommandButton CBCuves 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les CUVES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   4860
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   3060
               Width           =   1935
            End
            Begin VB.CommandButton CBMatieres 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les matières"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   300
               Style           =   1  'Graphical
               TabIndex        =   26
               Top             =   4020
               Width           =   2055
            End
            Begin VB.CommandButton CBGammesProduction 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les GAMMES de PRODUCTION"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   300
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   4860
               Width           =   4275
            End
            Begin VB.CommandButton CBActionsPossibles 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les ACTIONS POSSIBLES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   4860
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   3960
               Width           =   6495
            End
            Begin VB.CommandButton CBBains 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les BAINS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   6000
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   2220
               Width           =   1935
            End
            Begin VB.CommandButton CBPonts 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les PONTS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   9420
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   3060
               Width           =   1935
            End
            Begin VB.CommandButton CBPostes 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les POSTES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   7140
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   3060
               Width           =   1935
            End
            Begin VB.CommandButton CBZones 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les ZONES de la ligne"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   20
               Top             =   480
               Width           =   7095
            End
            Begin VB.CommandButton CBTempsMouvements 
               BackColor       =   &H00FFFFC0&
               Caption         =   "Les TEMPS de MOUVEMENTS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   4860
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   4860
               Width           =   6495
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   1
               Left            =   3480
               TabIndex        =   48
               Top             =   4620
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":6625B
               maskpicture     =   "FOrganisationLigne.frx":6646D
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   2
               Left            =   5760
               TabIndex        =   49
               Top             =   2820
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":6667F
               maskpicture     =   "FOrganisationLigne.frx":66891
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   3
               Left            =   5760
               TabIndex        =   50
               Top             =   3720
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":66AA3
               maskpicture     =   "FOrganisationLigne.frx":66CB5
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   4
               Left            =   6300
               TabIndex        =   51
               Top             =   2820
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":66EC7
               maskpicture     =   "FOrganisationLigne.frx":670D9
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   5
               Left            =   8040
               TabIndex        =   52
               Top             =   2820
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":672EB
               maskpicture     =   "FOrganisationLigne.frx":674FD
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   6
               Left            =   8040
               TabIndex        =   53
               Top             =   3720
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":6770F
               maskpicture     =   "FOrganisationLigne.frx":67921
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   7
               Left            =   8040
               TabIndex        =   54
               Top             =   4620
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":67B33
               maskpicture     =   "FOrganisationLigne.frx":67D45
               maskcolor       =   16711935
            End
            Begin Anodisation.ImageMask IMFlecheVersLeBas 
               Height          =   240
               Index           =   8
               Left            =   10320
               TabIndex        =   55
               Top             =   3720
               Width           =   135
               _extentx        =   238
               _extenty        =   423
               picture         =   "FOrganisationLigne.frx":67F57
               maskpicture     =   "FOrganisationLigne.frx":68169
               maskcolor       =   16711935
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   35
               X1              =   8100
               X2              =   8100
               Y1              =   3480
               Y2              =   3720
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   34
               X1              =   5820
               X2              =   5820
               Y1              =   840
               Y2              =   2880
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   33
               X1              =   6360
               X2              =   6360
               Y1              =   2820
               Y2              =   2640
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   32
               X1              =   8100
               X2              =   8100
               Y1              =   840
               Y2              =   2820
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   5
               X1              =   1320
               X2              =   1320
               Y1              =   4380
               Y2              =   4620
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   4
               X1              =   5820
               X2              =   5820
               Y1              =   3480
               Y2              =   3720
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   0
               X1              =   10380
               X2              =   10380
               Y1              =   3420
               Y2              =   3720
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   9
               X1              =   3540
               X2              =   3540
               Y1              =   4620
               Y2              =   840
            End
            Begin VB.Line LRaccords 
               BorderColor     =   &H0000FFFF&
               BorderWidth     =   3
               Index           =   15
               X1              =   8100
               X2              =   8100
               Y1              =   4320
               Y2              =   4680
            End
         End
         Begin VB.PictureBox PBDonnees 
            BackColor       =   &H00C0E0FF&
            Height          =   12075
            Left            =   12060
            ScaleHeight     =   12015
            ScaleWidth      =   16455
            TabIndex        =   16
            Top             =   480
            Width           =   16515
            Begin TrueOleDBGrid80.TDBGrid TDBGDonnees 
               Height          =   1215
               Left            =   240
               TabIndex        =   72
               Top             =   120
               Width           =   4095
               _ExtentX        =   7223
               _ExtentY        =   2143
               _LayoutType     =   0
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   2
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   -1  'True
               Splits(0)._GSX_SAVERECORDSELECTORS=   0
               Splits(0).DividerColor=   14215660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=2"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
               Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
               Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   0
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               MultipleLines   =   0
               CellTipsWidth   =   0
               DeadAreaBackColor=   14215660
               RowDividerColor =   14215660
               RowSubDividerColor=   14215660
               DirectionAfterEnter=   1
               DirectionAfterTab=   1
               MaxRows         =   250000
               ViewColumnCaptionWidth=   0
               ViewColumnWidth =   0
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
               _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
               _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
               _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
               _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
               _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
               _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
               _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
               _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
               _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
               _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
               _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
               _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
               _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
               _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
               _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
               _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
               _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
               _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
               _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
               _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
               _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
               _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
               _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
               _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
               _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
               _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
               _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
               _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
               _StyleDefs(38)  =   "Named:id=33:Normal"
               _StyleDefs(39)  =   ":id=33,.parent=0"
               _StyleDefs(40)  =   "Named:id=34:Heading"
               _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(42)  =   ":id=34,.wraptext=-1"
               _StyleDefs(43)  =   "Named:id=35:Footing"
               _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(45)  =   "Named:id=36:Selected"
               _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(47)  =   "Named:id=37:Caption"
               _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(49)  =   "Named:id=38:HighlightRow"
               _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(51)  =   "Named:id=39:EvenRow"
               _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(53)  =   "Named:id=40:OddRow"
               _StyleDefs(54)  =   ":id=40,.parent=33"
               _StyleDefs(55)  =   "Named:id=41:RecordSelector"
               _StyleDefs(56)  =   ":id=41,.parent=34"
               _StyleDefs(57)  =   "Named:id=42:FilterBar"
               _StyleDefs(58)  =   ":id=42,.parent=33"
            End
            Begin RichTextLib.RichTextBox RTBDonnees 
               Height          =   615
               Left            =   180
               TabIndex        =   17
               Top             =   1560
               Visible         =   0   'False
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   1085
               _Version        =   393217
               Enabled         =   -1  'True
               ScrollBars      =   3
               TextRTF         =   $"FOrganisationLigne.frx":6837B
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
         End
         Begin VB.Label LTitreDonnees 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Left            =   12060
            TabIndex        =   46
            Top             =   180
            Width           =   16515
         End
         Begin VB.Label LTitres 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Logigramme de fonctionnement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Index           =   1
            Left            =   180
            TabIndex        =   45
            Top             =   6360
            Width           =   11715
         End
         Begin VB.Label LTitres 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Caractéristiques de la ligne"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   44
            Top             =   180
            Width           =   11715
         End
      End
   End
End
Attribute VB_Name = "FOrganisationLigne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant l'organisation de la ligne
' Nom                    : FOrganisationLigne.frm
' Date de création : 08/10/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const EPAISSEUR_LIGNE As Integer = 230
Private Const TITRE_FENETRE As String = "ORGANISATION DE LA LIGNE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum SOURCES_DONNEES    'les sources de données
    
    T_ZONES = 0
    
    T_MATIERES = 1
    T_GAMMES_ANODISATION = 2
    
    T_BAINS = 3
    T_CUVES = 4
    T_POSTES = 5
    T_PONTS = 6
    T_ACTIONS_POSSIBLES = 7
    T_TEMPS_MOUVEMENTS = 8
    
    T_CHARGEMENT = 9
    
    T_CYCLE_ACTUEL_PONT_1 = 10
    T_FIN_CYCLE_ACTUEL_PONT_1 = 11
    T_TRANSFERT_PROCHAIN_CYCLE_PONT_1 = 12
    T_PROCHAIN_CYCLE_PONT_1 = 13
    
    T_CYCLE_ACTUEL_PONT_2 = 14
    T_FIN_CYCLE_ACTUEL_PONT_2 = 15
    T_TRANSFERT_PROCHAIN_CYCLE_PONT_2 = 16
    T_PROCHAIN_CYCLE_PONT_2 = 17
    
    T_ANALYSE_CHARGES_EN_LIGNE = 18
    T_PREMISSES = 19
    T_IA_MOTEUR_INFERENCE = 20

    T_PREVISIONNEL = 21

End Enum

Private Enum COLONNES_MATIERES

    C_ORDRE_POUR_AFFICHAGE = 0
    C_MATIERE = 1
    C_TYPE_MATIERE = 2
    C_COMPOSITION_MATIERE = 3

End Enum

Private Enum COLONNES_ZONES
    
    C_NUM_ZONE = 0
    C_CODE_ZONE = 1
    C_LIBELLE_ZONE = 2
    C_NUM_PREMIER_POSTE = 3
    C_NOM_PREMIER_POSTE = 4
    C_NUM_DERNIER_POSTE = 5
    C_NOM_DERNIER_POSTE = 6
    C_NBR_POSTES = 7
    
End Enum

Private Enum COLONNES_POSTES
    
    C_NUM_POSTE = 0
    C_NOM_POSTE = 1
    C_LIBELLE_POSTE = 2
    
    C_AVEC_TEMPS = 3
    C_RESPECT_TEMPS_OBLIGATOIRE = 4
    
    C_AVEC_EGOUTTAGE = 5
    
    C_COUVERCLES = 6
    C_REDRESSEUR = 7
    C_AGITATION_BAIN = 8
    
    C_X_AXE_POSTE_LIGNE = 9
    C_X_AXE_POSTE_SYNOPTIQUE = 10
    
    C_X_INFERIEUR_POSTE_SYNOPTIQUE = 11
    C_Y_INFERIEUR_POSTE_SYNOPTIQUE = 12
    C_X_SUPERIEUR_POSTE_SYNOPTIQUE = 13
    C_Y_SUPERIEUR_POSTE_SYNOPTIQUE = 14
    
    C_X_INFERIEUR_LIBELLE_POSTE_SYNOPTIQUE = 15
    C_Y_INFERIEUR_LIBELLE_POSTE_SYNOPTIQUE = 16
    C_X_SUPERIEUR_LIBELLE_POSTE_SYNOPTIQUE = 17
    C_Y_SUPERIEUR_LIBELLE_POSTE_SYNOPTIQUE = 18

End Enum

Private Enum COLONNES_CUVES
    
    C_NUM_CUVE = 0
    C_NOM_CUVE = 1
    C_LIBELLE_CUVE = 2
    
    C_GESTION_API = 3
    
    C_PRESENCE_POMPE = 4
    C_NBR_CHAUFFAGES = 5
    C_PRESENCE_REFROIDISSEMENT_BAIN = 6
    C_PRESENCE_NIVEAU_BAS = 7
    C_PRESENCE_NIVEAU_HAUT = 8
    C_PRESENCE_EV_EAU = 9
    C_PRESENCE_ANALYSEUR = 10

End Enum

Private Enum COLONNES_BAINS
    
    C_NUM_BAIN = 0
    C_NOM_BAIN = 1
    C_LIBELLE_BAIN = 2
    
End Enum

Private Enum COLONNES_ACTIONS
    
    C_NUM_ACTION = 0
    C_CODE_ACTION = 1
    C_LIBELLE_ACTION = 2
    C_PARAMATRE_OUI_NON = 3
    C_LIBELLE_PARAMETRE = 4
    
End Enum

Private Enum COLONNES_PONTS
    
    C_NUM_PONT = 0
    C_NOM_PONT = 1
    C_LIBELLE_PONT = 2
    
End Enum

'--- types privés ---

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                                  'pour interdire certains évènements
Private LigneDepartDeplacement As Integer                              'ligne de départ en cas de déplacement d'un détail
Private LigneArriveeDeplacement As Integer                             'ligne de d'arrivée en cas de déplacement d'un détail
Private NbrColonnesDonnees As Integer                                   'nombre de colonnes des données
Private NbrLignesDonnees As Long                                           'nombre de lignes des données
Private MemDernierBouton As Long                                           'mémoire du dernier bouton
Private NbrEnregistrementsEnCours As Long                            'indique le nombre d'enregistrement en cours
Private NomsChampsEnCours As String                                     'noms des champs en cours
Private RequeteEnCours As String                                              'indique la requête en cours
Private SourceDonneesEnCours As SOURCES_DONNEES        'source de données en cours de traitement

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBActionsPossibles_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_ACTIONS_POSSIBLES
End Sub

Private Sub CBActualiser_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- gestion des boutons ---
    GestionBoutons E_AVANT_ACTUALISER
    
    '--- curseur de la souris ---
    SourisEnAttente True
    
    '--- actualisation ---
    AiguillageSourcesDonnees SourceDonneesEnCours

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

Private Sub CBAnalyseChargesEnCours_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_ANALYSE_CHARGES_EN_LIGNE
End Sub

Private Sub CBChargement_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_CHARGEMENT
End Sub

Private Sub CBCuves_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_CUVES
End Sub

Private Sub CBCycleActuelPont_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine d'aiguillage ---
    Select Case Index
        Case PONTS.P_1: AiguillageSourcesDonnees T_CYCLE_ACTUEL_PONT_1
        Case PONTS.P_2: AiguillageSourcesDonnees T_CYCLE_ACTUEL_PONT_2
        Case Else
    End Select

End Sub

Private Sub CBEnregistrements_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- analyse en fonction du bouton appuyé ---
    Select Case Index
        
        Case BOUTONS_ENREGISTREMENTS.B_PREMIER
            
            TDBGDonnees.MoveFirst
            

        Case BOUTONS_ENREGISTREMENTS.B_PRECEDENT
            '--- précédent ---
            If Not TDBGDonnees.BOF Then
                TDBGDonnees.MovePrevious
            End If
        
        Case BOUTONS_ENREGISTREMENTS.B_SUIVANT
            '--- suivant ---
            If Not TDBGDonnees.EOF Then
                TDBGDonnees.MoveNext
            End If
        
        Case BOUTONS_ENREGISTREMENTS.B_DERNIER
            '--- fin ---
            TDBGDonnees.MoveLast
        
        Case Else
    
    End Select
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
      
    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number

End Sub

Private Sub CBEnregistrements_GotFocus(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déplacement du focus sur le bouton ---
    With SFocus
        .Left = PBBoutonsEnregistrements.Left
        .Top = PBBoutonsEnregistrements.Top
        .Height = PBBoutonsEnregistrements.Height
        .Width = PBBoutonsEnregistrements.Width
        .Visible = True
    End With

End Sub

Private Sub CBEnregistrements_LostFocus(Index As Integer)
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBGammesProduction_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_GAMMES_ANODISATION
End Sub

Private Sub CBMoteurInference_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_IA_MOTEUR_INFERENCE
End Sub

Private Sub CBPonts_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_PONTS
End Sub

Private Sub CBPostes_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_POSTES
End Sub

Private Sub CBPremisses_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_PREMISSES
End Sub

Private Sub CBPrevisionnel_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_PREVISIONNEL
End Sub

Private Sub CBProchainCyclePont_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine d'aiguillage ---
    Select Case Index
        Case PONTS.P_1: AiguillageSourcesDonnees T_PROCHAIN_CYCLE_PONT_1
        Case PONTS.P_2: AiguillageSourcesDonnees T_PROCHAIN_CYCLE_PONT_2
        Case Else
    End Select
    
End Sub

Private Sub CBQuitter_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rechargement des postes pour modifier le tableau commun
    '    notamment pour le respect des temps obligatoires ---
    Bidon = ChargePostes()
    
    '--- déchargement de la fenêtre ---
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

Private Sub CBTempsMouvements_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_TEMPS_MOUVEMENTS
End Sub

Private Sub CBTransfertProchainCycleEnActuel_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine d'aiguillage ---
    Select Case Index
        Case PONTS.P_1: AiguillageSourcesDonnees T_TRANSFERT_PROCHAIN_CYCLE_PONT_1
        Case PONTS.P_2: AiguillageSourcesDonnees T_TRANSFERT_PROCHAIN_CYCLE_PONT_2
        Case Else
    End Select

End Sub

Private Sub CBBains_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_BAINS
End Sub

Private Sub CBMatieres_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_MATIERES
End Sub

Private Sub CBZones_Click()
    On Error Resume Next
    AiguillageSourcesDonnees T_ZONES
End Sub

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        PremiereActivation = True
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Aiguillage des sources de données
' Entrées : SourceDonneesSouhaitee -> Source de données souhaitée fonction
'                                                                de l'énumération SOURCES_DONNEES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AiguillageSourcesDonnees(ByVal SourceDonneesSouhaitee As SOURCES_DONNEES)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumPont As Integer

    '--- affichage du titre ou aiguillage vers les écrans concernés ---
    Select Case SourceDonneesSouhaitee
        
        Case SOURCES_DONNEES.T_ZONES
            '--- les zones ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les ZONES"
            Bidon = GestionZones(GG_INITIALISATION)
            If GestionZones(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionZones(GG_AFFICHAGE)
            End If
        
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
    
        Case SOURCES_DONNEES.T_MATIERES
            '--- les matières ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les MATIERES"
            Bidon = GestionMatieres(GG_INITIALISATION)
            If GestionMatieres(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionMatieres(GG_AFFICHAGE)
            End If
            
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
        
        Case SOURCES_DONNEES.T_GAMMES_ANODISATION
            '--- les gammes d'anodisation ---
            AppelFenetre F_GAMMES_ANODISATION
        
        Case SOURCES_DONNEES.T_BAINS
            '--- les bains ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les BAINS"
            Bidon = GestionBains(GG_INITIALISATION)
            If GestionBains(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionBains(GG_AFFICHAGE)
            End If
            
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
            
        Case SOURCES_DONNEES.T_CUVES
            '--- les cuves ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les CUVES"
            Bidon = GestionCuves(GG_INITIALISATION)
            If GestionCuves(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionCuves(GG_AFFICHAGE)
            End If
            
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
            
        Case SOURCES_DONNEES.T_POSTES
            '--- les postes ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les POSTES"
            Bidon = GestionPostes(GG_INITIALISATION)
            If GestionPostes(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionPostes(GG_AFFICHAGE)
            End If
            
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
        
        Case SOURCES_DONNEES.T_PONTS
            '--- les ponts ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les PONTS"
            Bidon = GestionPonts(GG_INITIALISATION)
            If GestionPonts(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionPonts(GG_AFFICHAGE)
            End If
            
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
            
        Case SOURCES_DONNEES.T_ACTIONS_POSSIBLES
            '--- les actions possibles ---
            TDBGDonnees.Visible = False   'effacement des grilles par défaut
            RTBDonnees.Visible = False
            
            '--- construction de la grille ---
            LTitreDonnees.Caption = "Les ACTIONS POSSIBLES"
            Bidon = GestionActions(GG_INITIALISATION)
            If GestionActions(GG_TRANSFERT_DONNEES) = "" Then
                Bidon = GestionActions(GG_AFFICHAGE)
            End If
            
            '--- mémorisation de la source de données en cours pour certaines fonctions ---
            SourceDonneesEnCours = SourceDonneesSouhaitee
            
        Case SOURCES_DONNEES.T_TEMPS_MOUVEMENTS
            '--- les temps de mouvements ---
            AppelFenetre FENETRES.F_TEMPS_MOUVEMENTS
        
        Case SOURCES_DONNEES.T_CHARGEMENT
            '--- chargement ---
            AppelFenetre FENETRES.F_CHARGEMENT_PREVISIONNEL
        
        Case SOURCES_DONNEES.T_CYCLE_ACTUEL_PONT_1, _
                 SOURCES_DONNEES.T_CYCLE_ACTUEL_PONT_2
            '--- affichage des cycles actuels sur la fenêtre du synoptique ---
            NumPont = IIf(SourceDonneesSouhaitee = SOURCES_DONNEES.T_CYCLE_ACTUEL_PONT_1, PONTS.P_1, PONTS.P_2)
            TEtatsPonts(NumPont).TypesAffichagesCyclesPonts = False       'FALSE=forçage du cycle actuel
            
            '--- sélection de la grille ---
            Select Case NumPont
                Case PONTS.P_1: AppelFenetre FENETRES.F_CYCLES_PONTS, FORMES_CYCLES_PONTS.F_CYCLES_PONT_1
                Case PONTS.P_2: AppelFenetre FENETRES.F_CYCLES_PONTS, FORMES_CYCLES_PONTS.F_CYCLES_PONT_2
                Case Else
            End Select
        
        Case SOURCES_DONNEES.T_FIN_CYCLE_ACTUEL_PONT_1, _
                 SOURCES_DONNEES.T_FIN_CYCLE_ACTUEL_PONT_2
            '--- fins des cycles actuels des ponts 1 et 2 ---
        
        Case SOURCES_DONNEES.T_TRANSFERT_PROCHAIN_CYCLE_PONT_1, _
                 SOURCES_DONNEES.T_TRANSFERT_PROCHAIN_CYCLE_PONT_2
            '--- transferts des prochains cycles des ponts 1 et 2 ---
        
        Case SOURCES_DONNEES.T_PROCHAIN_CYCLE_PONT_1, _
                 SOURCES_DONNEES.T_PROCHAIN_CYCLE_PONT_2
            '--- affichage des cycles actuels sur la fenêtre du synoptique ---
            NumPont = IIf(SourceDonneesSouhaitee = SOURCES_DONNEES.T_PROCHAIN_CYCLE_PONT_1, PONTS.P_1, PONTS.P_2)
            TEtatsPonts(NumPont).TypesAffichagesCyclesPonts = True        'TRUE=forçage du prochain cycle
            
            '--- sélection de la grille ---
            Select Case NumPont
                Case PONTS.P_1: AppelFenetre FENETRES.F_CYCLES_PONTS, FORMES_CYCLES_PONTS.F_CYCLES_PONT_1
                Case PONTS.P_2: AppelFenetre FENETRES.F_CYCLES_PONTS, FORMES_CYCLES_PONTS.F_CYCLES_PONT_2
                Case Else
            End Select
    
        Case SOURCES_DONNEES.T_ANALYSE_CHARGES_EN_LIGNE
            '--- appel de l'écran des charges en ligne ---
            AppelFenetre F_CHARGES_EN_LIGNE
       
        Case SOURCES_DONNEES.T_PREMISSES
            '--- appel de l'écran des prémisses ---
            AppelFenetre FENETRES.F_PREMISSES
        
        Case SOURCES_DONNEES.T_IA_MOTEUR_INFERENCE
            '--- le moteur d'inférence ---
            AppelFenetre FENETRES.F_MOTEUR_INFERENCE
        
        Case SOURCES_DONNEES.T_PREVISIONNEL
            '--- prévisionnel ---
            AppelFenetre FENETRES.F_CHARGEMENT_PREVISIONNEL, 1
        
        Case Else

    End Select

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
    
    '--- déclaration ---
    
    '--- zone mére et fille du déplacement de la fenetre ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_MERE).Height = Abs(Me.ScaleHeight - PBBoutons.Height)
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

Private Sub IMFinCycleActuelPont_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine d'aiguillage ---
    Select Case Index
        Case PONTS.P_1: AiguillageSourcesDonnees T_FIN_CYCLE_ACTUEL_PONT_1
        Case PONTS.P_2: AiguillageSourcesDonnees T_FIN_CYCLE_ACTUEL_PONT_2
        Case Else
    End Select

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
    PBBoutonsEnregistrements.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - PBBoutonsEnregistrements.Width
    CBActualiser.Left = PBBoutonsEnregistrements.Left - MARGES.M_ENTRE_BOUTONS - CBActualiser.Width

    '--- ensemble des grilles de données ---
    With TDBGDonnees
        .Left = PBDonnees.ScaleLeft
        .Top = PBDonnees.ScaleTop
        .Width = PBDonnees.ScaleWidth
        .Height = PBDonnees.ScaleHeight
    End With
    With RTBDonnees
        .Left = PBDonnees.ScaleLeft
        .Top = PBDonnees.ScaleTop
        .Width = PBDonnees.ScaleWidth
        .Height = PBDonnees.ScaleHeight
    End With

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
' Rôle      : Gére l'états des boutons des enregistrements après une action de l'opérateur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionBoutonsEnregistrements(ByVal Situation As ETATS_BOUTONS_ENREGISTREMENTS)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim UnBouton As CommandButton
    
    Select Case Situation
    
        Case ETATS_BOUTONS_ENREGISTREMENTS.E_TOUT_INVISIBLE
            '--- rendre tous les boutons invisibles ---
            PBBoutonsEnregistrements.Visible = False
            For Each UnBouton In CBEnregistrements
                UnBouton.Visible = False
            Next
        
        Case ETATS_BOUTONS_ENREGISTREMENTS.E_TOUT_VISIBLE
            '--- rendre tous les boutons visibles ---
            PBBoutonsEnregistrements.Visible = True
            For Each UnBouton In CBEnregistrements
                UnBouton.Visible = True
            Next
        
        Case ETATS_BOUTONS_ENREGISTREMENTS.E_PRECEDENT_SUIVANT
            '--- rendre les boutons suivant et précédent visibles, verrouillé les autres ---
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_PREMIER)
                .Enabled = False
                .Visible = True
            End With
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_PRECEDENT)
                .Enabled = True
                .Visible = True
            End With
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_SUIVANT)
                .Enabled = True
                .Visible = True
            End With
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_DERNIER)
                .Enabled = False
                .Visible = True
            End With
        
        Case ETATS_BOUTONS_ENREGISTREMENTS.E_PREMIER_DERNIER
            '--- rendre les boutons début et fin visibles, verrouillé les autres ---
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_PREMIER)
                .Enabled = True
                .Visible = True
            End With
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_PRECEDENT)
                .Enabled = False
                .Visible = True
            End With
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_SUIVANT)
                .Enabled = False
                .Visible = True
            End With
            With CBEnregistrements(BOUTONS_ENREGISTREMENTS.B_DERNIER)
                .Enabled = True
                .Visible = True
            End With
        
        Case Else

    End Select

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
            CBActualiser.Enabled = True
        
        Case ETATS_BOUTONS.E_DECHARGEMENT_FENETRE
            '--- au déchargement de la fenêtre ---
        
        Case ETATS_BOUTONS.E_AVANT_VALIDER
            '--- avant valider ---
        
        Case ETATS_BOUTONS.E_APRES_VALIDER
            '--- après valider ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ANNULER
            '--- avant annuler ---
        
        Case ETATS_BOUTONS.E_APRES_ANNULER
            '--- après annuler ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True

        Case ETATS_BOUTONS.E_AVANT_ACTUALISER
            '--- avant actualiser ---
        
        Case ETATS_BOUTONS.E_APRES_ACTUALISER
            '--- après actualiser ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
        
        Case ETATS_BOUTONS.E_MODIFICATION_EN_COURS
            '--- après modifier (à ne pas traiter si nouvel enregistrement) ---
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
            CBQuitter.Enabled = True
            CBActualiser.Enabled = False

        Case ETATS_BOUTONS.E_AVANT_NOUVEAU
            '--- avant nouveau ---
        
        Case ETATS_BOUTONS.E_APRES_NOUVEAU
            '--- après nouveau ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = False
        
        Case ETATS_BOUTONS.E_AVANT_SUPPRIMER
            '--- avant supprimer ---
        
        Case ETATS_BOUTONS.E_APRES_SUPPRIMER
            '--- après supprimer ---
            CBQuitter.Enabled = True
            CBActualiser.Enabled = True
        
        Case Else
    
    End Select

    '--- affectation ---
    MemDernierBouton = Situation

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
' Rôle      : Gestion des postes
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionPostes(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---

    '--- déclaration ---
    Dim NomCommande As String

    '--- affectation ---
    GestionPostes = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'division de la grille
            
                .HeadLines = 4                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = False                                                  'interdire un nouvel enregistrement
                .AllowDelete = False                                                     'interdire la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire la sélection des colonnes
                .AllowColMove = False                                                 'interdire le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Postes"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_POSTES.C_NUM_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du poste"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_NOM_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du poste"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_LIBELLE_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé du poste"
                    .Width = EPAISSEUR_CARACTERE * 25
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_AVEC_TEMPS)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Avec temps (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_POSTES.C_RESPECT_TEMPS_OBLIGATOIRE)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Respect du temps obligatoire (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_POSTES.C_AVEC_EGOUTTAGE)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Avec égouttage (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_POSTES.C_COUVERCLES)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Couvercles (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_POSTES.C_REDRESSEUR)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Redresseur (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_POSTES.C_AGITATION_BAIN)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Agitation du bain (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_POSTES.C_X_AXE_POSTE_LIGNE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "X" & vbCr & "Axe de poste" & vbCr & "LIGNE (laser)"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_X_AXE_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "X" & vbCr & "Axe de poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_X_INFERIEUR_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "X Inférieur" & vbCr & "Poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_Y_INFERIEUR_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Y Inférieur" & vbCr & "Poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_X_SUPERIEUR_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "X Supérieur" & vbCr & "Poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_Y_SUPERIEUR_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Y Supérieur" & vbCr & "Poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_X_INFERIEUR_LIBELLE_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "X Inférieur" & vbCr & "Libellé du poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_Y_INFERIEUR_LIBELLE_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Y Inférieur" & vbCr & "Libellé du poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_X_SUPERIEUR_LIBELLE_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "X Supérieur" & vbCr & "Libellé du poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_POSTES.C_Y_SUPERIEUR_LIBELLE_POSTE_SYNOPTIQUE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Y Supérieur" & vbCr & "Libellé du poste" & vbCr & "SYNOPTIQUE"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With

                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionPostes = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des cuves
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionCuves(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomCommande As String
    
    '--- affectation ---
    GestionCuves = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure

                .Splits(0).AllowSizing = True                                        'division de la grille
            
                .HeadLines = 4                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = False                                                  'interdire un nouvel enregistrement
                .AllowDelete = False                                                     'interdire la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire la sélection des colonnes
                .AllowColMove = False                                                 'interdire le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Cuves"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_CUVES.C_NUM_CUVE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° de cuve"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_CUVES.C_NOM_CUVE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom de la cuve"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_CUVES.C_LIBELLE_CUVE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé de la cuve"
                    .Width = EPAISSEUR_CARACTERE * 25
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_CUVES.C_GESTION_API)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Gestion par l'automate (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_CUVES.C_PRESENCE_POMPE)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Pompe (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_CUVES.C_NBR_CHAUFFAGES)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nombre de chauffages"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgRight
                End With
                
                With .Columns(COLONNES_CUVES.C_PRESENCE_REFROIDISSEMENT_BAIN)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Groupe froid (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_CUVES.C_PRESENCE_NIVEAU_BAS)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Niveau très bas (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_CUVES.C_PRESENCE_NIVEAU_HAUT)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Niveau très haut (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_CUVES.C_PRESENCE_EV_EAU)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Arrivée d'eau (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_CUVES.C_PRESENCE_ANALYSEUR)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Analyseur (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionCuves = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des actions
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionActions(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomCommande As String

    '--- affectation ---
    GestionActions = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'division de la grille
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = True                                                    'interdire ou autoriser un nouvel enregistrement
                .AllowDelete = True                                                       'interdire ou autoriser la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire ou autoriser la sélection des colonnes
                .AllowColMove = False                                                 'interdire  ou autoriser le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'interdire ou autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'interdire ou autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Actions"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_ACTIONS.C_NUM_ACTION)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° de l'action"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_ACTIONS.C_CODE_ACTION)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Code de l'action"
                    .Width = EPAISSEUR_CARACTERE * 13
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_ACTIONS.C_LIBELLE_ACTION)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé de l'action"
                    .Width = EPAISSEUR_CARACTERE * 50
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_ACTIONS.C_PARAMATRE_OUI_NON)
                    .Locked = False
                    .ValueItems.Presentation = dbgCheckBox
                    .Caption = "Paramètre (OUI/NON)"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_ACTIONS.C_LIBELLE_PARAMETRE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé du paramètre"
                    .Width = EPAISSEUR_CARACTERE * 50
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With

                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionActions = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des zones
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionZones(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomCommande As String
    
    '--- affectation ---
    GestionZones = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'division de la grille
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = True                                                    'interdire ou autoriser un nouvel enregistrement
                .AllowDelete = True                                                       'interdire ou autoriser la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire ou autoriser la sélection des colonnes
                .AllowColMove = False                                                 'interdire  ou autoriser le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'interdire ou autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'interdire ou autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Zones"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_ZONES.C_NUM_ZONE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° de la zone"
                    .Width = EPAISSEUR_CARACTERE * 7
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_ZONES.C_CODE_ZONE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Code de la zone"
                    .Width = EPAISSEUR_CARACTERE * 15
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_ZONES.C_LIBELLE_ZONE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé de la zone"
                    .Width = EPAISSEUR_CARACTERE * 40
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_ZONES.C_NUM_PREMIER_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du premier poste"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_ZONES.C_NOM_PREMIER_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du premier poste"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_ZONES.C_NUM_DERNIER_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du dernier poste"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_ZONES.C_NOM_DERNIER_POSTE)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du dernier poste"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With
                
                With .Columns(COLONNES_ZONES.C_NBR_POSTES)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nombre de postes"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgCenter
                End With

                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionZones = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des bains
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionBains(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomCommande As String
    
    '--- affectation ---
    GestionBains = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'division de la grille
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = True                                                   'interdire ou autoriser un nouvel enregistrement
                .AllowArrows = True                                                     'interdire ou autoriser la gestion des touches du clavier
                .AllowDelete = True                                                      'interdire ou autoriser la suppression d'un nouvel enregistrement
                .AllowUpdate = True                                                     'interdire ou autoriser la mise à jour
                
                .AllowColSelect = False                                                'interdire ou autoriser la sélection des colonnes
                .AllowColMove = False                                                 'interdire ou autoriser le déplacement des colonnes sélectionnées
                
                .MultiSelect = dbgMultiSelectSimple                            'type de sélection
                .AllowRowSelect = True                                                 'interdire ou autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'interdire ou autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Bains"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_BAINS.C_NUM_BAIN)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du bain"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_BAINS.C_NOM_BAIN)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du bain"
                    .Width = EPAISSEUR_CARACTERE * 8
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_BAINS.C_LIBELLE_BAIN)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé du bain"
                    .Width = EPAISSEUR_CARACTERE * 25
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With

                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionBains = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des ponts
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionPonts(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomCommande As String
    
    '--- affectation ---
    GestionPonts = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure
            
                .Splits(0).AllowSizing = True                                        'division de la grille
            
                .HeadLines = 3                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
                
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = False                                                  'interdire un nouvel enregistrement
                .AllowDelete = False                                                     'interdire la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire la sélection des colonnes
                .AllowColMove = False                                                 'interdire le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Ponts"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_PONTS.C_NUM_PONT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "N° du pont"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_PONTS.C_NOM_PONT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Nom du pont"
                    .Width = EPAISSEUR_CARACTERE * 12
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_PONTS.C_LIBELLE_PONT)
                    .Locked = True
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Libellé du pont"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With

                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionPonts = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des matières
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
' Retours : "" indique aucun incident sinon le numéro de l'erreur est renvoyé
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GestionMatieres(ByVal EtatSouhaite As GESTION_GRILLES) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomCommande As String
    
    '--- affectation ---
    GestionMatieres = ""

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation de la grille ---
            With TDBGDonnees
                
                .Visible = False                                                            'rendre la grille invisible
                .ClearFields                                                                  'effacer la structure
                
                .Splits(0).AllowSizing = False                                       'division de la grille
            
                .HeadLines = 2                                                             'nombre de ligne des entêtes
                .HeadBackColor = COULEURS.ROUGE_5                   'couleur de fond des entêtes
                .HeadForeColor = COULEURS.BLANC                         'couleur de plan des entêtes
                
                .DeadAreaBackColor = COULEURS.ORANGE_0          'couleur de la surface non utilisée
      
                .AlternatingRowStyle = True                                         'couleur des lignes en alternance
                .EvenRowStyle.BackColor = COULEURS.VERT_1       'couleur des lignes paires
                .OddRowStyle.BackColor = COULEURS.JAUNE_1       'couleur des lignes impaires
                
                With .HeadFont
                    .Name = "Arial"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                With .Font
                    .Name = "MS Sans serif"
                    .Size = 10
                    .Bold = True                                                              'caractères gras
                End With
                
                .RowHeight = 0                                                              'épaisseur des lignes
                .RowHeight = .RowHeight * 1.06
                
                .RecordSelectors = True                                                'affichage du sélecteur d'enregistrement
                .RecordSelectorWidth = EPAISSEUR_CARACTERE * 3 'épaisseur du sélecteur d'enregistrement
                .RecordSelectorStyle.BackColor = .HeadBackColor      'couleur de fond du sélecteur d'enregistrement
                .RecordSelectorStyle.ForeColor = COULEURS.BLANC  '.HeadForeColor     'couleur de plan du sélecteur d'enregistrement
                
                .TransparentRowPictures = True
                Set .PictureCurrentRow = OccFPrincipale.ILGrillesDonnees.ListImages("fleche blanche").Picture
                Set .PictureModifiedRow = OccFPrincipale.ILGrillesDonnees.ListImages("modification blanche").Picture
                Set .PictureAddnewRow = OccFPrincipale.ILGrillesDonnees.ListImages("etoile blanche").Picture
        
                .AllowAddNew = True                                                    'interdire ou autoriser un nouvel enregistrement
                .AllowDelete = True                                                       'interdire ou autoriser la suppression d'un nouvel enregistrement
                
                .AllowColSelect = False                                                'interdire ou autoriser la sélection des colonnes
                .AllowColMove = False                                                 'interdire  ou autoriser le déplacement des colonnes sélectionnées
                
                .AllowRowSelect = True                                                 'interdire ou autoriser la sélection des lignes
                .AllowRowSizing = True                                                 'interdire ou autoriser la modification de l'épaisseur des lignes
                
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données vers la grille ---
            NomCommande = Me.Name & "_Matieres"
            With TDBGDonnees
                Set .DataSource = ED
                .DataMember = NomCommande
                .ReOpen
            End With

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With TDBGDonnees
                
                With .Columns(COLONNES_MATIERES.C_ORDRE_POUR_AFFICHAGE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Ordre d'affichage"
                    .Width = EPAISSEUR_CARACTERE * 10
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_MATIERES.C_MATIERE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Matière"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_MATIERES.C_TYPE_MATIERE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Type"
                    .Width = EPAISSEUR_CARACTERE * 30
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                With .Columns(COLONNES_MATIERES.C_COMPOSITION_MATIERE)
                    .Locked = False
                    .ValueItems.Presentation = dbgNormal
                    .Caption = "Composition de la matière"
                    .Width = EPAISSEUR_CARACTERE * 50
                    .HeadingStyle.Alignment = dbgCenter
                    .HeadingStyle.TransparentForegroundPicture = True
                    .HeadingStyle.ForegroundPicturePosition = dbgFPLeftOfText
                    .HeadingStyle.ForegroundPicture = OccFPrincipale.ILGrillesDonnees.ListImages(IIf(.Locked = True, "indicateur rouge", "indicateur vert")).Picture
                    .Alignment = dbgGeneral
                End With
                
                .Visible = True
            
            End With

        Case Else

    End Select
    
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
    GestionMatieres = CStr(Err.Number)

End Function

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
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- fermeture de l'enregistrement et de la connexion ---
    
    '--- sauvegarde de certaines valeurs dans la base des registres ---
    
    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFOrganisationLigne = Nothing
    
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
        .WindowState = vbMaximized
    End With
    
    '--- images des fonds ---
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondOrange2
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- gestion de l'états des boutons ---
    GestionBoutons E_CHARGEMENT_FENETRE

    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:

    '--- affichage du message d'erreur ---
    MessageErreur TITRE_MESSAGES, Err.Description, Err.Number
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Afficher des données au format RTF
' Entrées : CouleurDonnees -> Couleur du texte du dialogue
'                    TexteDonnees -> Texte des données à afficher
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheDonnees(ByVal CouleurDonnees As COULEURS, _
                                              ByVal NbrTabulations As Integer, _
                                              ByVal TexteDonnees As String, _
                                              ByVal NbrNouvellesLignes As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim a As Integer
    Dim TexteFormate As String
    
    With Me.RTBDonnees
        
        '--- changement de couleur ---
        .SelColor = CouleurDonnees
        
        '--- formatage du texte ---
        For a = 1 To NbrTabulations
            TexteFormate = TexteFormate & vbTab
        Next a
        TexteFormate = TexteFormate & TexteDonnees
        For a = 1 To NbrNouvellesLignes
            TexteFormate = TexteFormate & vbCrLf
        Next a
        
        '--- affichage du texte ---
        .SelText = TexteFormate
        .SelStart = Len(.Text)
 
        .Refresh
    
        '--- forcer la couleur bleu par défaut ---
        .SelColor = COULEURS.BLEU_3
    
    End With

End Sub

Private Sub TDBGDonnees_Error(ByVal DataError As Integer, Response As Integer)
    On Error Resume Next
    Response = vbDataErrContinue
End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.Value
End Sub

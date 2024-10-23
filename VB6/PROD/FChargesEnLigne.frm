VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FChargesEnLigne 
   ClientHeight    =   14145
   ClientLeft      =   525
   ClientTop       =   2415
   ClientWidth     =   28680
   Icon            =   "FChargesEnLigne.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   943
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1912
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FChargesEnLigne.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   28620
      TabIndex        =   2
      Top             =   0
      Width           =   28680
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
         Left            =   300
         TabIndex        =   3
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
      Picture         =   "FChargesEnLigne.frx":24A8C
      ScaleHeight     =   1035
      ScaleWidth      =   28620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   13050
      Width           =   28680
      Begin VB.PictureBox PBRedresseurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Index           =   4
         Left            =   5220
         Picture         =   "FChargesEnLigne.frx":25BB2
         ScaleHeight     =   885
         ScaleWidth      =   2505
         TabIndex        =   135
         Top             =   60
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Label LTempsRestantCycle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   142
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label LURedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   4
            Left            =   1365
            TabIndex        =   138
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label LIRedresseurS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   4
            Left            =   1365
            TabIndex        =   137
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label LModeRedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   136
            Top             =   90
            Width           =   1155
         End
      End
      Begin VB.PictureBox PBRedresseurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Index           =   3
         Left            =   9780
         Picture         =   "FChargesEnLigne.frx":26313
         ScaleHeight     =   885
         ScaleWidth      =   2505
         TabIndex        =   131
         Top             =   60
         Width           =   2535
         Begin VB.Label LTempsRestantCycle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   141
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label LURedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   3
            Left            =   1365
            TabIndex        =   134
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label LIRedresseurS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   3
            Left            =   1365
            TabIndex        =   133
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label LModeRedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   132
            Top             =   90
            Width           =   1155
         End
      End
      Begin VB.PictureBox PBRedresseurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Index           =   2
         Left            =   13080
         Picture         =   "FChargesEnLigne.frx":26A74
         ScaleHeight     =   885
         ScaleWidth      =   2505
         TabIndex        =   127
         Top             =   60
         Width           =   2535
         Begin VB.Label LTempsRestantCycle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   140
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label LURedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   2
            Left            =   1365
            TabIndex        =   130
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label LIRedresseurS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   2
            Left            =   1365
            TabIndex        =   129
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label LModeRedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   128
            Top             =   90
            Width           =   1155
         End
      End
      Begin VB.PictureBox PBRedresseurs 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   915
         Index           =   1
         Left            =   16380
         Picture         =   "FChargesEnLigne.frx":271D5
         ScaleHeight     =   885
         ScaleWidth      =   2505
         TabIndex        =   123
         Top             =   60
         Width           =   2535
         Begin VB.Label LTempsRestantCycle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   139
            Top             =   480
            Width           =   1155
         End
         Begin VB.Label LURedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   1
            Left            =   1365
            TabIndex        =   126
            Top             =   90
            Width           =   1035
         End
         Begin VB.Label LIRedresseurS 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   1
            Left            =   1365
            TabIndex        =   125
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label LModeRedresseurs 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            ForeColor       =   &H00000080&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   124
            Top             =   90
            Width           =   1155
         End
      End
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   116
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   30
            Left            =   0
            SmallChange     =   10
            TabIndex        =   119
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   30
            Left            =   900
            SmallChange     =   10
            TabIndex        =   118
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FChargesEnLigne.frx":27936
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
            Picture         =   "FChargesEnLigne.frx":27AE0
            Style           =   1  'Graphical
            TabIndex        =   117
            ToolTipText     =   " Agrandissement de la fenêtre "
            Top             =   0
            UseMaskColor    =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Timer TimerSynoptique 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   3360
         Top             =   60
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FChargesEnLigne.frx":27C8A
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
         Left            =   24960
         MaskColor       =   &H00FF00FF&
         Picture         =   "FChargesEnLigne.frx":2838C
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBReduire 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Réduire la fenêtre"
         DownPicture     =   "FChargesEnLigne.frx":28A8E
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
         Left            =   22680
         MaskColor       =   &H00FF00FF&
         Picture         =   "FChargesEnLigne.frx":29190
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Réduire cette fenêtre à la taille minimum "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2115
      End
      Begin MSComctlLib.ImageList ILImagesPourGrilles 
         Left            =   1380
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   10
         ImageHeight     =   10
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargesEnLigne.frx":29892
               Key             =   "indicateur bleu"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargesEnLigne.frx":29A2A
               Key             =   "indicateur rouge"
            EndProperty
         EndProperty
      End
      Begin VB.Timer TimerChargeGeree 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   2820
         Top             =   60
      End
      Begin MSComctlLib.ImageList ILOutilsSynoptique 
         Left            =   2100
         Top             =   60
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   39
         ImageHeight     =   19
         MaskColor       =   16711935
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargesEnLigne.frx":29BBC
               Key             =   "croix de condamnation"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargesEnLigne.frx":2A4F6
               Key             =   "rectangle rouge"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargesEnLigne.frx":2AE30
               Key             =   "rectangle blanc"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FChargesEnLigne.frx":2B76A
               Key             =   "rectangle vert"
            EndProperty
         EndProperty
      End
      Begin VB.Label LNumPhaseEnCours 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   4
         Left            =   4620
         TabIndex        =   146
         Top             =   375
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label LNumPhaseEnCours 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   9180
         TabIndex        =   145
         Top             =   375
         Width           =   615
      End
      Begin VB.Label LNumPhaseEnCours 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   12480
         TabIndex        =   144
         Top             =   360
         Width           =   615
      End
      Begin VB.Label LNumPhaseEnCours 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   1
         Left            =   15780
         TabIndex        =   143
         Top             =   360
         Width           =   615
      End
      Begin VB.Image IDecoration 
         Height          =   915
         Index           =   6
         Left            =   4620
         Picture         =   "FChargesEnLigne.frx":2C0A4
         Top             =   60
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image IDecoration 
         Height          =   915
         Index           =   5
         Left            =   9180
         Picture         =   "FChargesEnLigne.frx":2DE72
         Top             =   60
         Width           =   615
      End
      Begin VB.Image IDecoration 
         Height          =   915
         Index           =   4
         Left            =   12480
         Picture         =   "FChargesEnLigne.frx":2FC40
         Top             =   60
         Width           =   615
      End
      Begin VB.Image IDecoration 
         Height          =   915
         Index           =   3
         Left            =   15780
         Picture         =   "FChargesEnLigne.frx":31A0E
         Top             =   60
         Width           =   615
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   3960
         Top             =   60
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12975
      Index           =   0
      Left            =   0
      ScaleHeight     =   865
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1912
      TabIndex        =   1
      Top             =   375
      Width           =   28680
      Begin VB.PictureBox PBDeplacementFenetre 
         Height          =   12795
         Index           =   1
         Left            =   0
         ScaleHeight     =   849
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   1919
         TabIndex        =   4
         Top             =   0
         Width           =   28845
         Begin VB.PictureBox PBDateFinDansLePoste 
            BackColor       =   &H00FF0000&
            Height          =   435
            Left            =   0
            Picture         =   "FChargesEnLigne.frx":337DC
            ScaleHeight     =   375
            ScaleWidth      =   28665
            TabIndex        =   114
            Top             =   3000
            Width           =   28725
            Begin VB.Label LDateFinDansLePoste 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
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
               ForeColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   6720
               TabIndex        =   115
               Top             =   60
               Width           =   14835
            End
         End
         Begin VB.PictureBox PBImageLigne 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3000
            Left            =   270
            Picture         =   "FChargesEnLigne.frx":5811E
            ScaleHeight     =   200
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1877
            TabIndex        =   113
            Top             =   0
            Width           =   28155
         End
         Begin C1SizerLibCtl.C1Tab CTOnglets 
            Height          =   8835
            Left            =   180
            TabIndex        =   7
            Top             =   3660
            Width           =   28215
            _cx             =   49768
            _cy             =   15584
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
            Caption         =   "Détails de la charge|Gamme d'ANODISATION|Globalité des temps|Traçabilité de la charge|Alarmes de la ligne"
            Align           =   0
            CurrTab         =   1
            FirstTab        =   0
            Style           =   1
            Position        =   0
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   5
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
            Picture(0)      =   "FChargesEnLigne.frx":16B160
            Picture(1)      =   "FChargesEnLigne.frx":16B2BA
            Picture(2)      =   "FChargesEnLigne.frx":16B414
            Picture(3)      =   "FChargesEnLigne.frx":16B56E
            Picture(4)      =   "FChargesEnLigne.frx":16B6C8
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   9
               Left            =   30960
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   17
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   1
               Left            =   45
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   16
               Top             =   495
               Width           =   28125
               Begin VB.Frame FGammeAnodisation 
                  Caption         =   " Référence de la gamme "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3495
                  Left            =   18240
                  TabIndex        =   24
                  Top             =   120
                  Width           =   9675
                  Begin VB.CommandButton CBValiderNouveauPointeur 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "Validation du nouveau pointeur"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   675
                     Left            =   540
                     Style           =   1  'Graphical
                     TabIndex        =   149
                     Top             =   4860
                     Width           =   2895
                  End
                  Begin VB.TextBox TBNouveauPointeur 
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
                     Left            =   1620
                     TabIndex        =   148
                     Top             =   4380
                     Width           =   735
                  End
                  Begin VB.CommandButton CBModificationOptionsCharge 
                     BackColor       =   &H00C0E0FF&
                     Caption         =   "Modification des options de la charge"
                     DownPicture     =   "FChargesEnLigne.frx":16B822
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
                     Left            =   5340
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargesEnLigne.frx":16CC44
                     Style           =   1  'Graphical
                     TabIndex        =   73
                     Top             =   2580
                     UseMaskColor    =   -1  'True
                     Width           =   4155
                  End
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
                     Left            =   2460
                     MultiLine       =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   71
                     Top             =   1680
                     Width           =   7035
                  End
                  Begin VB.ComboBox CBChoixPosteAnodisation 
                     CausesValidation=   0   'False
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
                     ItemData        =   "FChargesEnLigne.frx":16E066
                     Left            =   2460
                     List            =   "FChargesEnLigne.frx":16E079
                     Style           =   2  'Dropdown List
                     TabIndex        =   69
                     Top             =   2400
                     Width           =   2655
                  End
                  Begin VB.CommandButton CBRechercheGammeAnodisation 
                     Height          =   315
                     Left            =   3900
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargesEnLigne.frx":16E0BA
                     Style           =   1  'Graphical
                     TabIndex        =   26
                     ToolTipText     =   " Lancer une recherche "
                     Top             =   420
                     UseMaskColor    =   -1  'True
                     Width           =   315
                  End
                  Begin VB.CommandButton CBExtensionReferenceGamme 
                     BackColor       =   &H00C0E0FF&
                     Height          =   315
                     Left            =   3420
                     MaskColor       =   &H00FF00FF&
                     Picture         =   "FChargesEnLigne.frx":16E3FC
                     Style           =   1  'Graphical
                     TabIndex        =   25
                     ToolTipText     =   " Modification du pointeur de la gamme "
                     Top             =   2940
                     UseMaskColor    =   -1  'True
                     Width           =   495
                  End
                  Begin VB.Line LNouveauPointeur 
                     BorderColor     =   &H000000FF&
                     BorderWidth     =   2
                     Index           =   1
                     Visible         =   0   'False
                     X1              =   1980
                     X2              =   1980
                     Y1              =   3420
                     Y2              =   3780
                  End
                  Begin VB.Line LNouveauPointeur 
                     BorderColor     =   &H000000FF&
                     BorderWidth     =   2
                     Index           =   0
                     Visible         =   0   'False
                     X1              =   540
                     X2              =   3420
                     Y1              =   3420
                     Y2              =   3420
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     Appearance      =   0  'Flat
                     BackColor       =   &H00C0FFFF&
                     BorderStyle     =   1  'Fixed Single
                     Caption         =   "MODIFICATION DU POINTEUR"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   9
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   255
                     Index           =   36
                     Left            =   540
                     TabIndex        =   147
                     Top             =   3960
                     Width           =   2895
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
                     Index           =   34
                     Left            =   540
                     TabIndex        =   122
                     Top             =   1320
                     Width           =   1740
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "Nom"
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
                     Index           =   33
                     Left            =   120
                     TabIndex        =   121
                     Top             =   900
                     Width           =   2160
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
                     Left            =   2460
                     TabIndex        =   120
                     Top             =   1260
                     Width           =   7050
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
                     Left            =   120
                     TabIndex        =   72
                     Top             =   1740
                     Width           =   2175
                     WordWrap        =   -1  'True
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
                     Left            =   2460
                     TabIndex        =   32
                     Top             =   840
                     Width           =   7050
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
                     Left            =   2460
                     TabIndex        =   31
                     Top             =   420
                     Width           =   1335
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     AutoSize        =   -1  'True
                     Caption         =   "GAMME N°"
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
                     Left            =   975
                     TabIndex        =   30
                     Top             =   480
                     Width           =   1320
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   2  'Center
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "POINTEUR de la ZONE en COURS"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Index           =   2
                     Left            =   660
                     TabIndex        =   29
                     Top             =   2880
                     Width           =   1665
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LPtrZoneGammeAnodisation 
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
                     Left            =   2460
                     TabIndex        =   28
                     Top             =   2940
                     Width           =   855
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Poste d'ANODISATION"
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
                     Index           =   1
                     Left            =   120
                     TabIndex        =   27
                     Top             =   2460
                     Width           =   2175
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Shape SNouveauPointeur 
                     BorderColor     =   &H000000FF&
                     BorderWidth     =   2
                     FillColor       =   &H00FFC0C0&
                     FillStyle       =   0  'Solid
                     Height          =   1965
                     Left            =   360
                     Top             =   3780
                     Visible         =   0   'False
                     Width           =   3255
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
                  Height          =   4335
                  Left            =   18240
                  TabIndex        =   76
                  Top             =   3720
                  Width           =   9675
                  Begin VB.PictureBox PBPhasesRedresseurs 
                     BackColor       =   &H00C0E0FF&
                     Height          =   3735
                     Left            =   3360
                     ScaleHeight     =   3675
                     ScaleWidth      =   6015
                     TabIndex        =   77
                     Top             =   360
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
                        TabIndex        =   79
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
                        TabIndex        =   83
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
                        TabIndex        =   89
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
                        TabIndex        =   95
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
                        TabIndex        =   80
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
                        TabIndex        =   85
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
                        TabIndex        =   91
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
                        TabIndex        =   97
                        Top             =   2460
                        Width           =   855
                     End
                     Begin MSMask.MaskEdBox MEBTempsPhases 
                        Height          =   315
                        Index           =   1
                        Left            =   1560
                        TabIndex        =   78
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
                        TabIndex        =   81
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
                        TabIndex        =   87
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
                        TabIndex        =   93
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
                        Index           =   35
                        Left            =   3840
                        TabIndex        =   106
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
                        Index           =   32
                        Left            =   3840
                        TabIndex        =   105
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
                        Index           =   31
                        Left            =   3840
                        TabIndex        =   104
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
                        Index           =   30
                        Left            =   3840
                        TabIndex        =   103
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
                        Index           =   29
                        Left            =   5400
                        TabIndex        =   102
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
                        Index           =   28
                        Left            =   2760
                        TabIndex        =   101
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
                        Index           =   27
                        Left            =   4320
                        TabIndex        =   100
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
                        Index           =   26
                        Left            =   5400
                        TabIndex        =   99
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
                        Index           =   25
                        Left            =   5400
                        TabIndex        =   98
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
                        Index           =   24
                        Left            =   5400
                        TabIndex        =   96
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
                        Index           =   19
                        Left            =   480
                        TabIndex        =   94
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
                        Index           =   20
                        Left            =   480
                        TabIndex        =   92
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
                        Index           =   21
                        Left            =   480
                        TabIndex        =   90
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
                        Index           =   16
                        Left            =   480
                        TabIndex        =   88
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
                        Index           =   13
                        Left            =   1440
                        TabIndex        =   86
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
                        TabIndex        =   84
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
                        Index           =   23
                        Left            =   480
                        TabIndex        =   82
                        Top             =   3000
                        Width           =   630
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
                        Index           =   6
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
                     Index           =   47
                     Left            =   225
                     TabIndex        =   110
                     Top             =   360
                     Width           =   2910
                  End
                  Begin VB.Image IPhasesAnodisation 
                     Height          =   2010
                     Left            =   240
                     Picture         =   "FChargesEnLigne.frx":16E73E
                     Top             =   1860
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
                     Index           =   48
                     Left            =   240
                     TabIndex        =   109
                     Top             =   1560
                     Width           =   2925
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
                     Left            =   600
                     TabIndex        =   108
                     Top             =   840
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
                     Left            =   1860
                     TabIndex        =   107
                     Top             =   840
                     Width           =   915
                  End
                  Begin VB.Shape SDecoration 
                     BorderWidth     =   2
                     FillColor       =   &H00FFFFC0&
                     FillStyle       =   0  'Solid
                     Height          =   675
                     Index           =   8
                     Left            =   240
                     Top             =   660
                     Width           =   2895
                  End
               End
               Begin MSMask.MaskEdBox MEBEditionDetailsGammesAnodisation 
                  Height          =   255
                  Left            =   420
                  TabIndex        =   33
                  Top             =   480
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
                  Height          =   7395
                  Left            =   240
                  TabIndex        =   34
                  Top             =   240
                  Width           =   17775
                  _ExtentX        =   31353
                  _ExtentY        =   13044
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
               Begin VB.Label LNumBarre 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FF80&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   405
                  Index           =   1
                  Left            =   9120
                  TabIndex        =   112
                  Top             =   7620
                  Width           =   8895
               End
               Begin VB.Label LNumCharge 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H0080FFFF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "-"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   405
                  Left            =   240
                  TabIndex        =   111
                  Top             =   7620
                  Width           =   8895
               End
               Begin VB.Shape SFocusTableDetailsGammesAnodisation 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   7800
                  Left            =   240
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   17790
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   0
               Left            =   -28770
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   15
               Top             =   495
               Width           =   28125
               Begin VB.Frame FRenseignements 
                  Caption         =   " Renseignements sur la charge "
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1215
                  Left            =   180
                  TabIndex        =   18
                  Top             =   120
                  Width           =   27735
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
                     Index           =   7
                     Left            =   6180
                     TabIndex        =   75
                     Top             =   360
                     Width           =   1395
                     WordWrap        =   -1  'True
                  End
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
                     Index           =   0
                     Left            =   7740
                     TabIndex        =   74
                     Top             =   300
                     Width           =   615
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
                     Height          =   300
                     Index           =   10
                     Left            =   120
                     TabIndex        =   22
                     Top             =   780
                     Width           =   3315
                     WordWrap        =   -1  'True
                  End
                  Begin VB.Label LChargePrioritaire 
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
                     TabIndex        =   21
                     Top             =   720
                     Width           =   585
                  End
                  Begin VB.Label LDateEntreeEnLigne 
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
                     TabIndex        =   20
                     Top             =   300
                     Width           =   2415
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
                     Index           =   12
                     Left            =   780
                     TabIndex        =   19
                     Top             =   360
                     Width           =   2655
                     WordWrap        =   -1  'True
                  End
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsCharges 
                  Height          =   6435
                  Left            =   240
                  TabIndex        =   23
                  Top             =   1560
                  Width           =   27615
                  _ExtentX        =   48710
                  _ExtentY        =   11351
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
               Begin VB.Shape SFocusTableDetailsCharges 
                  BorderColor     =   &H000000FF&
                  BorderWidth     =   4
                  Height          =   6450
                  Left            =   240
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   27630
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   3
               Left            =   29160
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   13
               Top             =   495
               Width           =   28125
               Begin VB.Frame FRenseignementsFicheProduction 
                  Caption         =   " Renseignements sur la fiche de traitement"
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
                  TabIndex        =   64
                  Top             =   120
                  Width           =   27735
                  Begin VB.Label LNbrPostesTraites 
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
                     Left            =   2640
                     TabIndex        =   66
                     Top             =   360
                     Width           =   1035
                  End
                  Begin VB.Label LLibelles 
                     Alignment       =   1  'Right Justify
                     Appearance      =   0  'Flat
                     AutoSize        =   -1  'True
                     BackColor       =   &H80000005&
                     BackStyle       =   0  'Transparent
                     Caption         =   "Nombre de postes traités"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H80000008&
                     Height          =   300
                     Index           =   3
                     Left            =   120
                     TabIndex        =   65
                     Top             =   360
                     Width           =   2385
                     WordWrap        =   -1  'True
                  End
               End
               Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFGDetailsFichesProduction 
                  Height          =   6810
                  Left            =   240
                  TabIndex        =   67
                  Top             =   1200
                  Width           =   27615
                  _ExtentX        =   48710
                  _ExtentY        =   12012
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
                  Height          =   6825
                  Left            =   240
                  Top             =   1200
                  Visible         =   0   'False
                  Width           =   27630
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   4
               Left            =   29460
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   12
               Top             =   495
               Width           =   28125
               Begin VB.TextBox TBAlarmesLigne 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   7815
                  Left            =   240
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  ScrollBars      =   3  'Both
                  TabIndex        =   68
                  Top             =   240
                  Width           =   27615
               End
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   5
               Left            =   29760
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   11
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   6
               Left            =   30060
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   10
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   7
               Left            =   30360
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   9
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   8
               Left            =   30660
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   8
               Top             =   495
               Width           =   28125
            End
            Begin VB.PictureBox PBOnglets 
               Height          =   8295
               Index           =   2
               Left            =   28860
               ScaleHeight     =   8235
               ScaleWidth      =   28065
               TabIndex        =   14
               Top             =   495
               Width           =   28125
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H00FFFFC0&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "PREVISIONS EN TEMPS REEL"
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
                  Index           =   11
                  Left            =   11220
                  TabIndex        =   70
                  Top             =   720
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LLibelles 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H000000FF&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "ATTENTION - Ces résultats théoriques peuvent varier de quelques dizaines de secondes avec la réalité"
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
                  Index           =   5
                  Left            =   180
                  TabIndex        =   63
                  Top             =   180
                  Width           =   27735
                  WordWrap        =   -1  'True
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
                  Index           =   6
                  Left            =   7980
                  TabIndex        =   62
                  Top             =   1380
                  Width           =   1815
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
                  Index           =   8
                  Left            =   3780
                  TabIndex        =   61
                  Top             =   1860
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
                  Left            =   3780
                  TabIndex        =   60
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
                  Left            =   3780
                  TabIndex        =   59
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
                  Left            =   3780
                  TabIndex        =   58
                  Top             =   3120
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
                  Left            =   3780
                  TabIndex        =   57
                  Top             =   2640
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
                  TabIndex        =   56
                  Top             =   720
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Left            =   11220
                  TabIndex        =   55
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   9135
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
                  TabIndex        =   54
                  Top             =   4140
                  Width           =   1035
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
                  TabIndex        =   53
                  Top             =   3660
                  Width           =   3195
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
                  TabIndex        =   52
                  Top             =   3180
                  Width           =   2535
                  WordWrap        =   -1  'True
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
                  TabIndex        =   51
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
                  TabIndex        =   50
                  Top             =   3600
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
                  TabIndex        =   49
                  Top             =   2640
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
                  TabIndex        =   48
                  Top             =   2700
                  Width           =   3315
                  WordWrap        =   -1  'True
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Left            =   11220
                  TabIndex        =   47
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   2
                  Left            =   11220
                  TabIndex        =   46
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   3
                  Left            =   11220
                  TabIndex        =   45
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   4
                  Left            =   11220
                  TabIndex        =   44
                  Top             =   2760
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   5
                  Left            =   11220
                  TabIndex        =   43
                  Top             =   3120
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   6
                  Left            =   11220
                  TabIndex        =   42
                  Top             =   3480
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   7
                  Left            =   11220
                  TabIndex        =   41
                  Top             =   3840
                  Visible         =   0   'False
                  Width           =   9135
               End
               Begin VB.Label LPrevisionsTempsReel 
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
                  Index           =   8
                  Left            =   11220
                  TabIndex        =   40
                  Top             =   4200
                  Visible         =   0   'False
                  Width           =   9135
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
                  TabIndex        =   39
                  Top             =   2640
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
                  TabIndex        =   38
                  Top             =   4080
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
                  TabIndex        =   37
                  Top             =   3600
                  Width           =   1815
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
                  Index           =   9
                  Left            =   5880
                  TabIndex        =   36
                  Top             =   1860
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
                  TabIndex        =   35
                  Top             =   3120
                  Width           =   1815
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
                  BorderWidth     =   4
                  Index           =   0
                  X1              =   5580
                  X2              =   6000
                  Y1              =   2820
                  Y2              =   2820
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
                  Index           =   2
                  X1              =   5580
                  X2              =   5880
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
                  Index           =   5
                  X1              =   7680
                  X2              =   7980
                  Y1              =   4260
                  Y2              =   4260
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
                  BorderStyle     =   3  'Dot
                  Index           =   8
                  X1              =   5250
                  X2              =   8430
                  Y1              =   3255
                  Y2              =   3255
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
                  Height          =   3675
                  Index           =   2
                  Left            =   10800
                  Shape           =   4  'Rounded Rectangle
                  Top             =   1080
                  Visible         =   0   'False
                  Width           =   9975
               End
            End
         End
         Begin VB.Image IDecoration 
            BorderStyle     =   1  'Fixed Single
            Height          =   3315
            Index           =   1
            Left            =   28440
            Picture         =   "FChargesEnLigne.frx":181B48
            Stretch         =   -1  'True
            Top             =   -195
            Width           =   300
         End
         Begin VB.Image IDecoration 
            BorderStyle     =   1  'Fixed Single
            Height          =   3315
            Index           =   0
            Left            =   0
            Picture         =   "FChargesEnLigne.frx":198ECA
            Stretch         =   -1  'True
            Top             =   -120
            Width           =   300
         End
      End
   End
End
Attribute VB_Name = "FChargesEnLigne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre gérant les charges en ligne
' Nom                    : FChargesEnLigne.frm
' Date de création : 09/12/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const NBR_COLONNES_DETAILS_CHARGES  As Integer = 7
Private Const NBR_COLONNES_DETAILS_GAMMES_ANODISATION  As Integer = 9
Private Const NBR_COLONNES_DETAILS_FICHES_PRODUCTION  As Integer = 7

Private Const TITRE_FENETRE As String = "CHARGES EN LIGNE"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum ONGLETS
    O_DETAILS_CHARGE = 0
    O_GAMME_ANODISATION = 1
    O_GLOBALITE_TEMPS = 2
    O_DETAILS_FICHE_PRODUCTION = 3
    O_ALARMES_LIGNE = 4
End Enum

Private Enum COLONNES_DETAILS_CHARGES
    C_NUM_LIGNES = 0
    C_NUM_COMMANDE_INTERNE = 1
    C_NBR_REPARATIONS = 2                         'nombre de réparations
    C_CODE_CLIENT = 3
    C_NBR_PIECES = 4
    C_DESIGNATION = 5
    C_NUM_LIGNES_REFERENCES_CLIENT = 6
    C_MATIERE = 7
End Enum

Private Enum COLONNES_DETAILS_GAMMES_ANODISATION
    C_NUM_LIGNES = 0
    C_CODE_ZONE = 1
    C_LIBELLE_ZONE = 2
    C_NOM_POSTE_REEL = 3
    C_TEMPS_AU_POSTE_TEXTE = 4
    C_DECOMPTE_TEMPS_POSTE_REEL = 5
    C_TEMPS_ALERTE_TEXTE = 6
    C_DECOMPTE_TEMPS_ALERTE = 7
    C_TEMPS_EGOUTTAGE_TEXTE = 8
    C_PONT = 9
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
Private Type ImgPriveeEtatsCharges
    DateEntreeEnLigne As Date                     'date d'entrée dans la ligne (généralement le chargement)
    ChargePrioritaire As Boolean                   'indique qu'il sagit  d'une charge prioritaire
    NumBarre As Integer                                'numéro de barre
    PtrZoneGammeAnodisation As Integer    'pointeur de la zone de la gamme d'anodisation
                                                                      'cette option est validé au chargement
    NbrPostesTraites As Integer                    'incrémentation de 1 à chaque entrée dans un poste
                                                                      'sert d'index pour les détails des fiches de production
    AlarmesLigne As String                            'alarmes de la ligne (séparation par le separateur de n° de défauts)
End Type

Private Type ImgDetailsCharges
    NumCommandeInterne As Long                           'n° de commande interne
    NbrReparations As String                                       'nombre de réparations (champ texte volontaire)
    CodeClient As String                                               'code client
    NbrPieces As String                                                'nombre de pièces
    Designation As String                                             'désignation
    NumLignesReferencesClient As String                  'n° de lignes des références du client correspondant
                                                                                    'aux n° de lignes des travaux avec les quantités séparés par des tirets
    NbrLignesReferencesClient As Integer                  'nombre de lignes des références du client une fois extraites
    Matiere As String                                                    'matière
End Type

Private Type ImgDetailsGammesAnodisation
    NumZone As Integer                                               'n° de la zone
    Codezone As String                                                'code de la zone
    LibelleZone As String                                             'libellé de la zone
    TempsAuPosteTexte As String                               'temps au poste en texte au format HH:MM:SS
    TempsAlerteTexte As String                                   'temps alerte en texte au format HH:MM:SS
    TempsEgouttageTexte As String                            'temps d'égouttage en texte au format MM:SS
    TempsAuPosteSecondes As Long                         'temps au poste en secondes
    TempsAlerteSecondes As Long                             'temps d'alerte en secondes
    TempsEgouttageSecondes As Integer                   'temps d'égouttage en secondes
    NomPosteReel As String                                        'nom du poste réel (cas des postes multiples)
    DecompteDuTempsAuPosteReelTexte As String   'décompte du temps au poste réel en texte (HH:MM:SS)
    DecompteDuTempsAlerteReelTexte As String       'décompte du temps d'alerte réel en texte (HH:MM:SS)
    FinDuTempsPosteReel As Boolean                       'TRUE = Indique la fin du temps au poste réel
    DebutAlertePosteReel As Boolean                         'TRUE = Indique le début de l'alerte au poste réel
End Type

Private Type ImgGammesAnodisation
    NumGamme As String                                                                                   'n° de gamme
    DateCreationGamme As Date                                                                        'date de création de la gamme
    NomGamme As String                                                                                   'nom de la gamme
    RefGamme As String                                                                                     'référence de la gamme
    Designation As String                                                                                    'désignation de la gamme d'anodisation
    TMatieresGamme(1 To NBR_MATIERES_MAXI_PAR_GAMME) As String    'tableau contenant les matières de la gamme
    TDetailsGammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesAnodisation
    ChoixPosteAnodisation As CHOIX_POSTE_ANODISATION                           'choix du poste d'anodisation
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
Private InterdireEvenements As Boolean                                      'pour interdire certains évènements

Private MemNumChargeEnCoursPourPhasesProduction As Integer     'mémoire de n° de charge en cours pour divers affichage
Private MemNumChargeEnCoursPourAffichage As Integer           'mémoire de n° de charge en cours pour divers affichage
Private MemNumChargeEnCoursPourAffichage1 As Integer         'mémoire de n° de charge en cours pour divers affichage
Private MemNumChargeEnCoursPourAffichage2 As Integer         'mémoire de n° de charge en cours pour divers affichage
Private NumChargeEnCours As Integer                                         'numéro de la charge en cours
Private MemNumLigne As Integer                                                 'mémoire d'un n° de ligne dans une des grilles
Private MemNumColonne As Integer                                             'mémoire d'un n° de colonne dans une des grilles

Private MemDernierBouton As Long                                              'mémoire du dernier bouton

Private ModeUouIEnCours As MODES_U_OU_I                            'mode U ou I en cours

'--- variables et tableaux privées DIRECTX 7.0 ---
Private ObjDX As New DirectX7                                                      'objet DirectX
Private ObjDD As DirectDraw7                                                        'objet DirectDraw
        
Private ObjDDSEcran As DirectDrawSurface7                                'objet de la surface de l'écran
Private DDSDEcran As DDSURFACEDESC2                                   'description de la surface de l'écran

Private ObjDDClip As DirectDrawClipper                                        'objet du clipper

Private ObjDDSImageLigne As DirectDrawSurface7                      'objet de la surface de l'image de la ligne
Private DDSDImageLigne As DDSURFACEDESC2                         'description de la surface de l'image de la ligne
Private RImageLigne As RECT                                                        'coordonnées du rectangle de l'image de la ligne

'--- tableaux privés ---
Private TPriveeEtatsCharges As ImgPriveeEtatsCharges
Private TDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As ImgDetailsCharges
Private TGammesAnodisation As ImgGammesAnodisation
Private TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As ImgDetailsPhasesProduction
Private TDetailsFichesProduction(1 To NBR_LIGNES_DETAILS_FICHES_PRODUCTION) As ImgDetailsFichesProduction

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBChoixPosteAnodisation_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
 
     If NumChargeEnCours > 0 And InterdireEvenements = False Then
                
        If TEtatsCharges(NumChargeEnCours).TGammesAnodisation.ChoixPosteAnodisation <> CBChoixPosteAnodisation.ListIndex Then
        
            If AppelFenetre(F_MESSAGE, _
                                    TITRE_MESSAGES, _
                                    vbCrLf & "cs|MODIFICATION DU POSTE d'anodisation" & vbCrLf & vbCrLf & _
                                    "Cette modification risque de bouleverser complètement" & vbCrLf & _
                                    "le fonctionnement de la ligne" & vbCrLf & vbCrLf & _
                                    "cs|Voulez-vous réellement valider ce changement ?", _
                                    TYPES_MESSAGES.T_ATTENTION, _
                                    TYPES_BOUTONS.T_OUI_NON, _
                                    EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
            
                '--- changer le poste d'anodisation directement dans les états des charges ---
                With TEtatsCharges(NumChargeEnCours).TGammesAnodisation
                    .ChoixPosteAnodisation = CBChoixPosteAnodisation.ListIndex
                End With
            
            Else
            
                '--- restaurer le bon poste d'anodisation dans l'outils ---
                InterdireEvenements = True
                CBChoixPosteAnodisation.ListIndex = TGammesAnodisation.ChoixPosteAnodisation
                InterdireEvenements = False
            
            End If

        End If
    
    End If

End Sub

Private Sub CBChoixPosteAnodisation_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    KeyCode = 0
End Sub

Private Sub CBChoixPosteAnodisation_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    KeyAscii = 0
End Sub

Private Sub CBExtensionReferenceGamme_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- modification de la dimension du cadre de la gamme ---
    ModificationDimensionsCadreGamme

End Sub

Private Sub CBModificationOptionsCharge_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la fenêtre de modification des options de la charge ---
    AppelFenetre FENETRES.F_MODIFICATION_OPTIONS_CHARGE, NumChargeEnCours

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

Private Sub CBRechercheGammeAnodisation_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la fenêtre des gammes d'anodisation ---
    'les paramêtres sont  TravailSurGrille
    '                                  RechercherPar
    '                                  CommencantPar
    '                                  Contenant
    '                                  MethodeRechercheChoisie
    If LNumGamme.Caption <> "" Then
        AppelFenetre FENETRES.F_GAMMES_ANODISATION, False, 1, LNumGamme.Caption, "", True
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

Private Sub CBValiderNouveauPointeur_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           NumPoste As Integer
    Dim PointeurZone As String

    '--- ouverture de la boite de dialogues et affichage ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
        
        '--- affectation du nouveau pointeur en texte / RAZ du champ d'édition ---
        PointeurZone = TBNouveauPointeur.Text
        TBNouveauPointeur.Text = ""
        
        If IsNumeric(PointeurZone) = True Then
            
            If CInt(PointeurZone) <= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                
                '--- changement du pointeur de la zone d'anodisation ---
                TEtatsCharges(NumChargeEnCours).PtrZoneGammeAnodisation = CInt(PointeurZone)
            
                '--- réinitialiser les lignes suivantes de la gamme ---
                With TEtatsCharges(NumChargeEnCours).TGammesAnodisation
                    For a = LBound(.TDetailsGammesAnodisation()) To UBound(.TDetailsGammesAnodisation())
                        If a > PointeurZone Then
                            With .TDetailsGammesAnodisation(a)
                                .NumPosteReel = 0                                               'N° de poste réel utilisé dans la zone (cas des postes multiples)
                                .DecompteDuTempsAuPosteReelSecondes = ""  'Représente la différence entre le temps théorique au poste
                                                                                                               'et le temps réel passé dans le poste
                                                                                                               'un nombre négatif apparait si la charge est resté plus
                                                                                                               'longtemps dans le poste que le temps théorique prévu
                                                                                                               'ATTENTION variable du type String volontairement
                                                                                                               'Si "" alors il n'y a pas eu de temps de décompter
                                .FinDuTempsPosteReel = False                            'TRUE = Indique la fin du temps au poste réel
                            End With
                        End If
                    Next a
                End With
                
                '--- enregistrement du n° de poste réel dans la gamme ---
                NumPoste = RechercheNumPostePourUneCharge(NumChargeEnCours)
                If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                    EnregistreNumPosteReelGamme NumPoste
                End If
            
            End If
        
        End If
    
    End If

End Sub

Private Sub CTOnglets_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    '--- focus ---
    Select Case CTOnglets.CurrTab

        Case ONGLETS.O_DETAILS_CHARGE
            '--- détails de la charge ---
            If MSHFGDetailsCharges.Visible = True Then MSHFGDetailsCharges.SetFocus

        Case ONGLETS.O_GAMME_ANODISATION
            '--- gamme Anodisation ---
            If MSHFGDetailsGammesAnodisation.Visible = True Then MSHFGDetailsGammesAnodisation.SetFocus

        Case ONGLETS.O_GLOBALITE_TEMPS
           '--- globalité des temps ---

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
    
    '--- déclaration ---
    Dim NumeroChargeLePlusPetit As Integer
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        
        '--- focus sur la grille des détails des charge ---
        MSHFGDetailsCharges.SetFocus
        
        '--- prendre le numéro de charge le plus petit comme charge en cours ---
        NumeroChargeLePlusPetit = RechercheNumeroChargeLePlusPetit()
        If NumeroChargeLePlusPetit > 0 Then
            NumChargeEnCours = NumeroChargeLePlusPetit
            EtatsChargeGeree
        End If
        
        '--- anti-rebond ---
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

Private Sub HSDeplacementFenetre_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Left = -HSDeplacementFenetre.value
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
    
        If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
        
            For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
        
                With TEtatsCharges(NumChargeEnCours).TGammesAnodisation.TDetailsGammesAnodisation(a)
                            
                    '--- recherche de la zone d'anodisation ---
                    If Trim(TZones(.NumZone).Codezone) = "C13 à C16" Then
                        
                        '--- affectation dans le tableau ---
                        .TempsAuPosteTexte = "0" & LTempsTotalGammeRedresseur.Caption
                        .TempsAuPosteSecondes = CTempsTexteEnSecondes(LTempsTotalGammeRedresseur.Caption)
                        
                        '--- rafraichissement dans la grille ---
                        MSHFGDetailsGammesAnodisation.TextMatrix(a, COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE) = .TempsAuPosteTexte
                        
                        Exit For
                    
                    End If
                                
                End With
        
            Next a

        End If

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

Private Sub MSHFGDetailsCharges_GotFocus()
    On Error Resume Next
    SFocusTableDetailsCharges.Visible = True
End Sub

Private Sub MSHFGDetailsCharges_LostFocus()
    On Error Resume Next
    SFocusTableDetailsCharges.Visible = False
End Sub

Private Sub MSHFGDetailsFichesProduction_GotFocus()
    On Error Resume Next
    SFocusTableDetailsFichesProduction.Visible = True
End Sub

Private Sub MSHFGDetailsFichesProduction_LostFocus()
    On Error Resume Next
    SFocusTableDetailsFichesProduction.Visible = False
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

Private Sub MSHFGDetailsGammesAnodisation_Scroll()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- rendre invisible le champ d'édition en cas de scrolling ---
    If MEBEditionDetailsGammesAnodisation.Visible = True Then
        MEBEditionDetailsGammesAnodisation.Visible = False
    End If

End Sub

Private Sub PBBoutons_Resize()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBReduire.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBReduire.Width
    
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
            If MemDernierBouton = ETATS_BOUTONS.E_APRES_NOUVEAU Then Exit Sub
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

Private Sub PBDateFinDansLePoste_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs des emplacements ---
    With PBDateFinDansLePoste
        LDateFinDansLePoste.Left = .ScaleLeft
        LDateFinDansLePoste.Top = .ScaleTop + 30
        LDateFinDansLePoste.Width = .ScaleWidth
        LDateFinDansLePoste.Height = .ScaleHeight
    End With

End Sub

Private Sub PBImageLigne_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer
    Dim NumPontClique  As Integer, _
           NumPosteClique As Integer, _
           NumLibellePosteClique As Integer, _
           NumCuve As Integer
    
    '--- recherche de la partie du synoptique cliqué pour les ponts ---
    For a = PONTS.P_1 To PONTS.P_2
        If X >= TXPonts(a) And Y >= TYPonts(a) And X <= (TXPonts(a) + DIMENSIONS_ANIMATIONS.D_LONG_PONT) And Y <= (TYPonts(a) + DIMENSIONS_ANIMATIONS.D_HAUT_PONT) Then
            NumPontClique = a
            Exit For
        End If
    Next a
    
    '--- recherche de la partie du synoptique cliqué pour les postes ---
    For a = POSTES.P_CHGT_1 To DERNIER_POSTE
        With TEtatsPostes(a).DefinitionPoste
            
            '--- recherche si poste cliqué ---
            If X >= .XInferieurPosteSynoptique And Y >= .YInferieurPosteSynoptique And X <= .XSuperieurPosteSynoptique And Y <= .YSuperieurPosteSynoptique Then
                NumPosteClique = a
                Exit For
            End If
            
            '--- recherche si libellé du poste cliqué ---
            If X >= .XInferieurLibellePosteSynoptique And Y >= .YInferieurLibellePosteSynoptique And X <= .XSuperieurLibellePosteSynoptique And Y <= .YSuperieurLibellePosteSynoptique Then
                NumLibellePosteClique = a
                Exit For
            End If
        
        End With
    Next a
    
    If Button = vbLeftButton Then
        
        '******************************************************************************************************
        '*                                                  ANALYSE SUR LE PONT CLIQUE
        '******************************************************************************************************
        If NumPontClique >= PONTS.P_1 And NumPontClique <= PONTS.P_2 Then
        
            With TEtatsPonts(NumPontClique)
        
                '--- affectation du numéro de charge ---
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    NumChargeEnCours = .NumCharge
                End If
        
            End With
        
        End If
        
        '******************************************************************************************************
        '*                                                 ANALYSE SUR LE POSTE CLIQUE
        '******************************************************************************************************
        If NumPosteClique >= POSTES.P_CHGT_1 And NumPosteClique <= DERNIER_POSTE Then
        
            With TEtatsPostes(NumPosteClique)
        
                '--- affectation du numéro de charge en cours ---
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    NumChargeEnCours = .NumCharge
                End If

            End With
        
        End If
        
        '******************************************************************************************************
        '*                                                 ANALYSE SUR LE LIBELLE CLIQUE
        '******************************************************************************************************
        If NumLibellePosteClique >= POSTES.P_CHGT_1 And NumLibellePosteClique <= DERNIER_POSTE Then
    
            With TEtatsPostes(NumLibellePosteClique)
            
                '--- affectation du numéro de charge en cours ---
                If .NumCharge >= CHARGES.C_NUM_MINI And .NumCharge <= CHARGES.C_NUM_MAXI Then
                    NumChargeEnCours = .NumCharge
                End If
            
            End With
        
        End If

    Else

        '--- condamnation du poste par un clic droit de la souris ---
        If NumLibellePosteClique >= POSTES.P_CHGT_1 And NumLibellePosteClique <= DERNIER_POSTE Then
            CondamnationPoste NumLibellePosteClique, TITRE_MESSAGES
        End If

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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées : NumCharge -> Numéro de charge souhaité
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre(ByVal NumCharge As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- numéro de charge en cours ---
    If NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI Then
        NumChargeEnCours = NumCharge
    End If
    
    '--- affichage ---
    AffichageChargeGeree
    
    GestionEtatsCharges GG_INITIALISATION
    GestionEtatsCharges GG_TRANSFERT_DONNEES
    GestionEtatsCharges GG_AFFICHAGE
    
    GestionDetailsCharges GG_INITIALISATION
    GestionDetailsCharges GG_TRANSFERT_DONNEES
    GestionDetailsCharges GG_AFFICHAGE
    
    GestionGammesAnodisation GG_INITIALISATION
    GestionGammesAnodisation GG_TRANSFERT_DONNEES
    GestionGammesAnodisation GG_AFFICHAGE
    
    GestionDetailsPhasesProduction GG_INITIALISATION
    GestionDetailsPhasesProduction GG_TRANSFERT_DONNEES
    GestionDetailsPhasesProduction GG_AFFICHAGE
    
    GestionDetailsFichesProduction GG_INITIALISATION
    GestionDetailsFichesProduction GG_TRANSFERT_DONNEES
    GestionDetailsFichesProduction GG_AFFICHAGE
    
    AffichageDateFinDansLePoste
    
    AffichageGlobaliteTemps

    '--- lancement des timers ---
    TimerChargeGeree.Enabled = True
    TimerSynoptique.Enabled = True

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
    
    '--- préparation de l'animation de la ligne ---
    InitialisationDirectX                          'initialisation de DirectX
    InitialisationSurfaces                        'Initialisation des surfaces
    
    '--- affectation ---
    CTOnglets.CurrTab = ONGLETS.O_GAMME_ANODISATION
    
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
    With TimerChargeGeree
        .Enabled = False
        .Interval = 0
    End With
    With TimerSynoptique
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFChargesEnLigne = Nothing

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
            PtrLigne As Integer, _
            NbrLignesReferencesClient As Integer
    Dim TempsEnSecondes As Double
    Dim FicheVide As ImgDetailsCharges, _
            TCopieDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As ImgDetailsCharges

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des détails ---
            Erase TDetailsCharges()

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
                
                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_CHARGES.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Numéro de pointage"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = "R."
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_CODE_CLIENT
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Code client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                .ColWidth(.Col) = 8 * EPAISSEUR_CARACTERE: .Text = "Nombre de pièces"
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_DESIGNATION
                .ColWidth(.Col) = 50 * EPAISSEUR_CARACTERE: .Text = "Désignation"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_NUM_LIGNES_REFERENCES_CLIENT
                .ColWidth(.Col) = 50 * EPAISSEUR_CARACTERE: .Text = "Quantité / référence du client"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_CHARGES.C_MATIERE
                .ColWidth(.Col) = 30 * EPAISSEUR_CARACTERE: .Text = "Matière"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                '--- centrage des titres ---
                .Row = 0
                For a = 1 To Pred(.Cols)
                    .Col = a
                    .CellAlignment = flexAlignCenterCenter
                Next a

                '--- N° de lignes, vidage des champs ---
                For a = LBound(TDetailsCharges()) To UBound(TDetailsCharges())
                
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
            For a = LBound(TDetailsCharges()) To UBound(TDetailsCharges())
                TDetailsCharges(a) = FicheVide
            Next a
            With MSHFGDetailsCharges
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- initialisation du tableau des détails ---
            Erase TDetailsCharges()

            '--- transfert des données dans le tableau ---
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                For a = LBound(TDetailsCharges()) To UBound(TDetailsCharges())
                    With TEtatsCharges(NumChargeEnCours).TDetailsCharges(a)
                        If .NumCommandeInterne = 0 Then
                            Exit For
                        Else
                            TDetailsCharges(a).NumCommandeInterne = .NumCommandeInterne
                            TDetailsCharges(a).NbrReparations = .NbrReparations
                            TDetailsCharges(a).CodeClient = .CodeClient
                            TDetailsCharges(a).NbrPieces = .NbrPieces
                            TDetailsCharges(a).Designation = .Designation
                            
                            '--- gestion des références du client ---
                            'TDetailsCharges(a).NumLignesReferencesClient = ExtraitReferencesClient(.NumCommandeInterne, _
                                                                                                                                                      .NumLignesReferencesClient, _
                                                                                                                                                      NbrLignesReferencesClient)
                            'TDetailsCharges(a).NbrLignesReferencesClient = NbrLignesReferencesClient
                            
                            'TDetailsCharges(a).Matiere = .Matiere
                        End If
                    End With
                Next a
            End If

        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- construction de la grille ---
            With MSHFGDetailsCharges

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TDetailsCharges()) To UBound(TDetailsCharges())
                    
                    .Row = a
                    
                    If TDetailsCharges(a).NumCommandeInterne = 0 Then
                        
                        TDetailsCharges(a) = FicheVide
                        For b = 1 To NBR_COLONNES_DETAILS_CHARGES
                            .Col = b
                            If .Text <> "" Then .Text = ""
                        Next b
                        .RowHeight(a) = .RowHeightMin
                    
                    Else
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_NUM_COMMANDE_INTERNE
                        AffichageTexte MSHFGDetailsCharges, TDetailsCharges(a).NumCommandeInterne

                        .Col = COLONNES_DETAILS_CHARGES.C_NBR_REPARATIONS
                        AffichageTexte MSHFGDetailsCharges, TDetailsCharges(a).NbrReparations
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_CODE_CLIENT
                        AffichageTexte MSHFGDetailsCharges, TDetailsCharges(a).CodeClient
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_NBR_PIECES
                        AffichageTexte MSHFGDetailsCharges, TDetailsCharges(a).NbrPieces
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_DESIGNATION
                        AffichageTexte MSHFGDetailsCharges, TDetailsCharges(a).Designation
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_NUM_LIGNES_REFERENCES_CLIENT
                        If .Text <> TDetailsCharges(a).NumLignesReferencesClient Then
                            .Text = TDetailsCharges(a).NumLignesReferencesClient
                            .RowHeight(a) = .RowHeightMin * 0.9 * TDetailsCharges(a).NbrLignesReferencesClient
                        End If
                        
                        .Col = COLONNES_DETAILS_CHARGES.C_MATIERE
                        AffichageTexte MSHFGDetailsCharges, TDetailsCharges(a).Matiere
                    
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
' Rôle      : Gestion des états des charges
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionEtatsCharges(ByVal EtatSouhaite As GESTION_GRILLES)
    
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
    Dim Texte As String
    Dim FicheVide As ImgPriveeEtatsCharges
    
    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau ---
            TPriveeEtatsCharges = FicheVide
            
            '--- vider les champs ---
            With LDateEntreeEnLigne
                .Caption = ""
                .Refresh
            End With
            With LChargePrioritaire
                .Caption = ""
                .Refresh
            End With
            With LNumBarre(0)
                .Caption = ""
                .Refresh
            End With
            With LNumBarre(1)
                .Caption = ""
                .Refresh
            End With
            With LPtrZoneGammeAnodisation
                .Caption = ""
                .Refresh
            End With
            With LNbrPostesTraites
                .Caption = ""
                .Refresh
            End With
            With TBAlarmesLigne
                .BackColor = COULEURS.BLANC
                .ForeColor = COULEURS.NOIR
                .Text = ""
                .Refresh
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- initialisation du tableau ---
            TPriveeEtatsCharges = FicheVide

            '--- transfert des données dans le tableau ---
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                With TEtatsCharges(NumChargeEnCours)
                    
                    '--- affectation ---
                    TPriveeEtatsCharges.DateEntreeEnLigne = .DateEntreeEnLigne
                    TPriveeEtatsCharges.ChargePrioritaire = .ChargePrioritaire
                    TPriveeEtatsCharges.NumBarre = .NumBarre
                    TPriveeEtatsCharges.PtrZoneGammeAnodisation = .PtrZoneGammeAnodisation
                    TPriveeEtatsCharges.NbrPostesTraites = .NbrPostesTraites
                    TPriveeEtatsCharges.AlarmesLigne = DecodeAlarmesLigne(.AlarmesLigne)
              
                End With
            End If
            
        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage ---
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                
                With TPriveeEtatsCharges
                
                    '--- affichage de la date d'entrée ---
                    If .DateEntreeEnLigne = Empty Then
                        Texte = ""
                    Else
                        Texte = UN_ESPACE & Format(.DateEntreeEnLigne, FORMAT_DATE_HEURE_1)
                    End If
                    AffichageTexte LDateEntreeEnLigne, Texte
                
                    '--- affichage indiquant si charge prioritaire ---
                    'si pas de date d'entrée valide alors afficher un blanc
                    If .DateEntreeEnLigne = Empty Then
                        Texte = ""
                    Else
                        Texte = UN_ESPACE & IIf(.ChargePrioritaire = True, "OUI", "NON")
                    End If
                    AffichageTexte LChargePrioritaire, Texte
                
                    '--- numéro de barre ---
                    If .NumBarre = 0 Then
                        Texte = "-"
                    Else
                        Texte = .NumBarre
                    End If
                    AffichageTexte LNumBarre(0), Texte
                    
                    '--- numéro de barre dans l'écran de la gamme ---
                    If .NumBarre = 0 Then
                        Texte = "-"
                    Else
                        Texte = "Barre n° " & .NumBarre
                    End If
                    AffichageTexte LNumBarre(1), Texte
                    
                    '--- affichage du pointeur de la zone ---
                    Texte = CStr(.PtrZoneGammeAnodisation)
                    AffichageTexte LPtrZoneGammeAnodisation, Texte
                     
                    '--- affichage du nombre de postes traités ---
                    Texte = CStr(.NbrPostesTraites)
                    AffichageTexte LNbrPostesTraites, Texte
                    
                    '--- affichage des alarmes de la ligne ---
                    Texte = .AlarmesLigne
                    With TBAlarmesLigne
                        If .Text <> Texte Then
                            If Texte = "" Then
                                .BackColor = COULEURS.BLANC
                                .ForeColor = COULEURS.NOIR
                            Else
                                .BackColor = COULEURS.ROUGE_3
                                .ForeColor = COULEURS.JAUNE_3
                            End If
                            .Text = Texte
                            .Refresh
                        End If
                    End With
                    
                End With
            
            End If

        Case Else

    End Select

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

Private Sub TBNouveauPointeur_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBNouveauPointeur_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 2
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

Private Sub Text1_Change()

End Sub

Private Sub TimerChargeGeree_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- appel de la routine ---
    TimerChargeGeree.Enabled = False
    EtatsChargeGeree
    AffichageDonneesRedresseurs
    TimerChargeGeree.Enabled = True
    
    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la date de fin dans le poste en cours
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageDateFinDansLePoste()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumPoste As Integer
    Dim TempsAuPosteSecondes As Long
    Dim Texte As String, _
            DecompteDuTempsAuPosteReelSecondes As String
    Dim DateFinDansLePoste As Date
            
    '--- affectation ---
    Texte = "-"
            
    '--- calcul de la date de sortie ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                
        With TEtatsCharges(NumChargeEnCours)
                    
            If .PtrZoneGammeAnodisation > 0 And .NbrPostesTraites > 0 Then
                    
                If .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel = .TDetailsFichesProduction(.NbrPostesTraites).NumPoste Then
                    
                    '--- affectation ---
                    NumPoste = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).NumPosteReel
                            
                    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                    
                        If TEtatsPostes(NumPoste).NumCharge = NumChargeEnCours Then
                    
                            '--- recherche du temps théorique dans la gamme ---
                            TempsAuPosteSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).TempsAuPosteSecondes
                            DecompteDuTempsAuPosteReelSecondes = .TGammesAnodisation.TDetailsGammesAnodisation(.PtrZoneGammeAnodisation).DecompteDuTempsAuPosteReelSecondes
                            
                            If TEtatsPostes(NumPoste).DefinitionPoste.AvecTemps = True Then
                                                        
                                With .TDetailsFichesProduction(.NbrPostesTraites)
                                               
                                    If .DateEntreePoste <> Empty Then
        
                                       '--- affectation ---
                                       DateFinDansLePoste = DateAdd("s", TempsAuPosteSecondes, .DateEntreePoste)
                                       Texte = "Sortie du poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                                                    " prévu le " & Format(DateFinDansLePoste, FORMAT_DATE_HEURE_1)
                                        If DecompteDuTempsAuPosteReelSecondes <> "" Then
                                            If IsNumeric(DecompteDuTempsAuPosteReelSecondes) = True Then
                                                Texte = Texte & ", temps restant = " & CTemps(CLng(DecompteDuTempsAuPosteReelSecondes))
                                            End If
                                        End If
                                    
                                    End If
        
                                End With
                                                                    
                            End If
                        
                        End If
                            
                    End If
                    
                End If
                    
            End If
                    
        End With
    
    End If

    '--- affichage du texte ---
    With LDateFinDansLePoste
        If .Caption <> Texte Then
            .Caption = Texte
            .Refresh
        End If
    End With
                
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la charge gérée
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageChargeGeree()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- contantes privées ---
    
    '--- déclaration ---
    Dim a As Integer
    Dim Texte As String
    Dim Texte1 As String

    '--- construction du texte ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
        
        '--- affectation ---
        Texte = "Charge n° " & NumChargeEnCours
        Texte1 = Texte
        
        '--- construction du texte avec les numéros de commandes internes ---
        For a = 1 To NBR_LIGNES_DETAILS_CHARGES
            With TEtatsCharges(NumChargeEnCours).TDetailsCharges(a)
                If .NumCommandeInterne = 0 Then
                    If a = 1 Then Texte = Texte & " - PAS DE REFERENCE"
                    Exit For
                Else
                    Texte = Texte & " - " & .NumCommandeInterne
                End If
            End With
        Next a
    
    Else
        
        '--- affectation ---
        Texte = "PAS DE CHARGE EN COURS"
        Texte1 = Texte
    
    End If

    '--- affichage ---
    With LRenseignementsFenetre
        If .Caption <> Texte Then
            .Caption = Texte
            .Refresh
        End If
    End With
    With LNumCharge
        If .Caption <> Texte1 Then
            .Caption = Texte1
            .Refresh
        End If
    End With

End Sub

Private Sub TimerSynoptique_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rafraichissement du synoptique ---
    TimerSynoptique.Enabled = False
    If OccFSynoptique.ArretTachesRapides = False Then
        GestionImageTampon
        TimerSynoptique.Enabled = True
    End If

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des gammes d'anodisation
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionGammesAnodisation(ByVal EtatSouhaite As GESTION_GRILLES)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const IMG_INDICATEUR_BLEU As String = "indicateur bleu"
    Const IMG_INDICATEUR_ROUGE As String = "indicateur rouge"
    
    '--- déclaration ---
    Dim TypeCouleur As Boolean
    Dim a As Integer, _
            b As Integer, _
            MemLigne As Integer, _
            MemColonne As Integer, _
            PtrLigne As Integer, _
            PtrZoneGammeAnodisation As Integer, _
            NumZoneDepart As Integer, _
            NumZoneArrivee As Integer, _
            NumPont As Integer
    Dim FicheVideGammesAnodisation As ImgGammesAnodisation, _
            FicheVideDetailsGammesAnodisation As ImgDetailsGammesAnodisation, _
            TCopieDetailsgammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As ImgDetailsGammesAnodisation

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION
            '--- initialisation du tableau des détails ---
           TGammesAnodisation = FicheVideGammesAnodisation

            '--- initialisation de la grille des détails ---
            With MSHFGDetailsGammesAnodisation

                .Redraw = False

                .Clear

                .FixedCols = 1
                .FixedRows = 1
                .Rows = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION + .FixedRows
                .Cols = NBR_COLONNES_DETAILS_GAMMES_ANODISATION + .FixedCols
                .RowSizingMode = flexRowSizeIndividual     'épaisseur de lignes modifiées ligne par ligne
                .RowHeight(0) = 820                                        'épaisseur des titres
                .RowHeightMin = 315
                .Row = 0

                '--- paramétrages de chaque colonne ---
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_NUM_LIGNES
                .ColWidth(.Col) = 3 * EPAISSEUR_CARACTERE: .Text = ""
                .ColAlignment(.Col) = flexAlignRightCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_CODE_ZONE
                .ColWidth(.Col) = 15 * EPAISSEUR_CARACTERE: .Text = "Code de la zone"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_LIBELLE_ZONE
                .ColWidth(.Col) = 36.5 * EPAISSEUR_CARACTERE: .Text = "Libellé de la zone"
                .ColAlignment(.Col) = flexAlignLeftCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_NOM_POSTE_REEL
                .ColWidth(.Col) = 10 * EPAISSEUR_CARACTERE: .Text = "Nom du poste"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Temps prévu au POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_DECOMPTE_TEMPS_POSTE_REEL
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Décompte du temps au POSTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Temps prévu d'ALERTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_DECOMPTE_TEMPS_ALERTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Décompte du temps d'ALERTE"
                .ColAlignment(.Col) = flexAlignCenterCenter
                
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                .ColWidth(.Col) = 11 * EPAISSEUR_CARACTERE: .Text = "Temps prévu d'égouttage"
                .ColAlignment(.Col) = flexAlignCenterCenter

                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_PONT
                .ColWidth(.Col) = 5 * EPAISSEUR_CARACTERE: .Text = "Pont"
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
                    .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_NUM_LIGNES
                    .Row = a
                    .Text = CStr(a)
                
                    '--- couleurs des lignes ---
                    .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_CODE_ZONE
                    .FillStyle = flexFillRepeat
                    .ColSel = COLONNES_DETAILS_GAMMES_ANODISATION.C_PONT
                    .CellBackColor = IIf(TypeCouleur = False, COULEURS.VERT_1, COULEURS.CYAN_1)
                    TypeCouleur = Not (TypeCouleur)
                
                Next a

                '--- fixer le curseur ---
                .Row = 1
                .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_CODE_ZONE

                .Redraw = True

            End With

        Case GESTION_GRILLES.GG_VIDAGE
            '--- vidage du tableau ---
            For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
                TGammesAnodisation.TDetailsGammesAnodisation(a) = FicheVideDetailsGammesAnodisation
            Next a
            With MSHFGDetailsGammesAnodisation
                .TopRow = 1
                .LeftCol = 1
            End With

        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- initialisation du tableau des détails ---
            TGammesAnodisation = FicheVideGammesAnodisation
            
            '--- transfert des données dans le tableau ---
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                With TEtatsCharges(NumChargeEnCours).TGammesAnodisation
                    
                    '--- N° et nom, etc ... ---
                    TGammesAnodisation.NumGamme = .NumGamme
                    TGammesAnodisation.DateCreationGamme = .DateCreationGamme
                    TGammesAnodisation.NomGamme = .NomGamme
                    TGammesAnodisation.RefGamme = .RefGamme
                    TGammesAnodisation.Designation = .Designation
                    For a = 1 To UBound(.TMatieresGamme())
                        TGammesAnodisation.TMatieresGamme(a) = .TMatieresGamme(a)
                    Next a
                    TGammesAnodisation.ChoixPosteAnodisation = .ChoixPosteAnodisation
                
                    '--- détails ---
                    For a = LBound(.TDetailsGammesAnodisation()) To UBound(.TDetailsGammesAnodisation())
                        With .TDetailsGammesAnodisation(a)
                            TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone = .NumZone
                            If .NumZone > 0 Then
                                
                                TGammesAnodisation.TDetailsGammesAnodisation(a).Codezone = TZones(.NumZone).Codezone
                                TGammesAnodisation.TDetailsGammesAnodisation(a).LibelleZone = TZones(.NumZone).LibelleZone
                                
                                TGammesAnodisation.TDetailsGammesAnodisation(a).TempsAuPosteTexte = .TempsAuPosteTexte
                                TGammesAnodisation.TDetailsGammesAnodisation(a).TempsAlerteTexte = .TempsAlerteTexte
                                TGammesAnodisation.TDetailsGammesAnodisation(a).TempsEgouttageTexte = .TempsEgouttageTexte
                                
                                TGammesAnodisation.TDetailsGammesAnodisation(a).TempsAuPosteSecondes = .TempsAuPosteSecondes
                                TGammesAnodisation.TDetailsGammesAnodisation(a).TempsAlerteSecondes = .TempsAlerteSecondes
                                TGammesAnodisation.TDetailsGammesAnodisation(a).TempsEgouttageSecondes = .TempsEgouttageSecondes
                                
                                '--- nom du poste réel (cas des postes multiples) ---
                                If .NumPosteReel >= POSTES.P_CHGT_1 And .NumPosteReel <= DERNIER_POSTE Then
                                    TGammesAnodisation.TDetailsGammesAnodisation(a).NomPosteReel = TEtatsPostes(.NumPosteReel).DefinitionPoste.NomPoste
                                Else
                                    TGammesAnodisation.TDetailsGammesAnodisation(a).NomPosteReel = ""
                                End If
                                
                                '--- décompte du temps réel au poste en HH:MM:SS ---
                                If .DecompteDuTempsAuPosteReelSecondes = "" Then
                                    TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAuPosteReelTexte = ""
                                Else
                                    If IsNumeric(.DecompteDuTempsAuPosteReelSecondes) = True Then
                                        TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAuPosteReelTexte = CTemps(CLng(.DecompteDuTempsAuPosteReelSecondes))
                                    Else
                                        TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAuPosteReelTexte = ""
                                    End If
                                End If
                                
                                '--- décompte du temps réel d'alerte en HH:MM:SS ---
                                If .DecompteDuTempsAlerteReelSecondes = "" Then
                                    TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAlerteReelTexte = ""
                                Else
                                    If IsNumeric(.DecompteDuTempsAlerteReelSecondes) = True Then
                                        TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAlerteReelTexte = CTemps(CLng(.DecompteDuTempsAlerteReelSecondes))
                                    Else
                                        TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAlerteReelTexte = ""
                                    End If
                                End If
                                
                                '--- fin du temps au poste réel ---
                                TGammesAnodisation.TDetailsGammesAnodisation(a).FinDuTempsPosteReel = .FinDuTempsPosteReel
                                
                                '--- début de l'alerte au poste réel ---
                                TGammesAnodisation.TDetailsGammesAnodisation(a).DebutAlertePosteReel = .DebutAlertePosteReel
                            
                            End If
                        End With
                    Next a
                
                End With
            End If
            
        Case GESTION_GRILLES.GG_COMPRESSION
            '--- compression des données ---
            PtrLigne = 1
            For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
                If TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone <> 0 Then
                    TCopieDetailsgammesAnodisation(PtrLigne) = TGammesAnodisation.TDetailsGammesAnodisation(a)
                    Inc PtrLigne
                End If
            Next a
            For a = 1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION
                TGammesAnodisation.TDetailsGammesAnodisation(a) = TCopieDetailsgammesAnodisation(a)
            Next a

        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage des caractéristiques de la gamme ---
            With TGammesAnodisation
          
                '--- n° de la gamme d'anodisation ---
                AffichageTexte LNumGamme, .NumGamme
          
                '--- nom de la gamme ---
                AffichageTexte LNomGamme, .NomGamme
                
                '--- référence de la gamme ---
                AffichageTexte LRefGamme, .RefGamme
          
                '--- matières concernées ---
                TBMatieresConcernees.Text = .TMatieresGamme(1)
                For a = 2 To UBound(.TMatieresGamme())
                    If .TMatieresGamme(a) <> "" Then
                        TBMatieresConcernees.Text = TBMatieresConcernees.Text & vbCrLf & .TMatieresGamme(a)
                    End If
                Next a
                
                '--- affichage du choix du poste d'anodisation ---
                InterdireEvenements = True
                If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                    If MemNumChargeEnCoursPourAffichage1 <> NumChargeEnCours Then
                        CBChoixPosteAnodisation.ListIndex = .ChoixPosteAnodisation
                        CBChoixPosteAnodisation.Locked = False
                        MemNumChargeEnCoursPourAffichage1 = NumChargeEnCours
                    End If
                Else
                    If CBChoixPosteAnodisation.Locked <> True Then
                        CBChoixPosteAnodisation.Locked = True
                    End If
                End If
                InterdireEvenements = False
            
            End With
            
            '--- construction de la grille ---
            With MSHFGDetailsGammesAnodisation

                '--- mémorisation des valeurs ligne, colonne ---
                MemLigne = .Row
                MemColonne = .Col
                .FocusRect = flexFocusNone
                .Redraw = False

                For a = LBound(TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(TGammesAnodisation.TDetailsGammesAnodisation())
                    
                    .Row = a
                    
                    If TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone = 0 Then
                        
                        TGammesAnodisation.TDetailsGammesAnodisation(a) = FicheVideDetailsGammesAnodisation
                        For b = 1 To NBR_COLONNES_DETAILS_GAMMES_ANODISATION
                            .Col = b
                            If .Text <> "" Then .Text = ""
                            If .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_DECOMPTE_TEMPS_POSTE_REEL Then
                                If .CellPicture <> LoadPicture() Then
                                    Set .CellPicture = LoadPicture()
                                End If
                            End If
                        Next b
                    
                    Else
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_CODE_ZONE
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).Codezone
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_LIBELLE_ZONE
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).LibelleZone
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_NOM_POSTE_REEL
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).NomPosteReel
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).TempsAuPosteTexte
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_DECOMPTE_TEMPS_POSTE_REEL
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAuPosteReelTexte
                       
                        '--- indicateur de fin de temps au poste ---
                        .CellPictureAlignment = flexAlignRightTop
                        If TGammesAnodisation.TDetailsGammesAnodisation(a).FinDuTempsPosteReel = False Then
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        Else
                            If .CellPicture <> Me.ILImagesPourGrilles.ListImages(IMG_INDICATEUR_BLEU).Picture Then
                                Set .CellPicture = Me.ILImagesPourGrilles.ListImages(IMG_INDICATEUR_BLEU).Picture
                            End If
                        End If
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).TempsAlerteTexte
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_DECOMPTE_TEMPS_ALERTE
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).DecompteDuTempsAlerteReelTexte
                        
                        '--- indicateur de début d'alerte au poste ---
                        .CellPictureAlignment = flexAlignRightTop
                        If TGammesAnodisation.TDetailsGammesAnodisation(a).DebutAlertePosteReel = False Then
                            If .CellPicture <> LoadPicture() Then
                                Set .CellPicture = LoadPicture()
                            End If
                        Else
                            If .CellPicture <> Me.ILImagesPourGrilles.ListImages(IMG_INDICATEUR_ROUGE).Picture Then
                                Set .CellPicture = Me.ILImagesPourGrilles.ListImages(IMG_INDICATEUR_ROUGE).Picture
                            End If
                        End If
                        
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                        AffichageTexte MSHFGDetailsGammesAnodisation, TGammesAnodisation.TDetailsGammesAnodisation(a).TempsEgouttageTexte
                    
                        '--- affectation des numéros de zones pour l'affichage du pont ---
                        NumZoneDepart = TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone
                        If a = NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                            NumZoneArrivee = 0
                        Else
                            If TGammesAnodisation.TDetailsGammesAnodisation(a + 1).Codezone = "" Then
                                NumZoneArrivee = 0
                            Else
                                NumZoneArrivee = TGammesAnodisation.TDetailsGammesAnodisation(a + 1).NumZone
                            End If
                        End If
                        
                        '--- affichage du pont ---
                        .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_PONT
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
                
                '--- rendre toujours visible indication du pointeur en cours ---
                If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                    If MemNumChargeEnCoursPourAffichage2 <> NumChargeEnCours Then
                        PtrZoneGammeAnodisation = TEtatsCharges(NumChargeEnCours).PtrZoneGammeAnodisation
                        If .RowIsVisible(PtrZoneGammeAnodisation) = False Then
                            If PtrZoneGammeAnodisation <= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                                .TopRow = PtrZoneGammeAnodisation
                            End If
                        End If
                        MemNumChargeEnCoursPourAffichage2 = NumChargeEnCours 'mémoire anti-rebond
                    End If
                End If

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
' Rôle      : Analyse la gamme pour afficher les redresseurs
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistreValeursRedresseursDansGamme()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    
    '--- déclaration ---
    
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
    
        '--- remplir dans le tableau des états des charges ---
        With TEtatsCharges(NumChargeEnCours)
                    
            '--- redresseur A3 ---
            'If PBRedresseurA3.Visible = True Then
            '    .TDetailsPhasesProduction(REDRESSEURS.R_C13).IProduction = CSng(TBIRedresseurs(IDX_REDRESSEURS.IDX_REDRESSEUR_A3).Text)
            '    If LAnodiqueCathodique.Caption = TEXTE_ANODIQUE Then 'signe négatif pour l'anodique
            '        .TDetailsPhasesProduction(REDRESSEURS.R_C13).IProduction = -.TDetailsPhasesProduction(REDRESSEURS.R_C13).IProduction
            '    End If
            'End If
            
            '--- redresseur A9 ---
            'If PBRedresseurA9.Visible = True Then
            '    .TDetailsPhasesProduction(REDRESSEURS.R_A9).IProduction = CSng(TBIRedresseurs(IDX_REDRESSEURS.IDX_REDRESSEUR_A9).Text)
            'End If
            
            '--- redresseur C13, C14, C15 ---
            'If PBRedresseurC13C14C15.Visible = True Then
                
                '--- pour C13 ---
                '.TDetailsPhasesProduction(REDRESSEURS.R_C13).UProduction = CSng(TBURedresseurs(IDX_REDRESSEURS.IDX_REDRESSEUR_C13_C14_C15).Text)
                '.TDetailsPhasesProduction(REDRESSEURS.R_C13).TempsAmorçageSecondes = CLng(TBTempsAmorcage.Text)
                
                '--- pour C14 ---
                '.TDetailsPhasesProduction(REDRESSEURS.R_C14).UProduction = CSng(TBURedresseurs(IDX_REDRESSEURS.IDX_REDRESSEUR_C13_C14_C15).Text)
                '.TDetailsPhasesProduction(REDRESSEURS.R_C14).TempsAmorçageSecondes = CLng(TBTempsAmorcage.Text)
                
                '--- pour C15 ---
                '.TDetailsPhasesProduction(REDRESSEURS.R_C15).UProduction = CSng(TBURedresseurs(IDX_REDRESSEURS.IDX_REDRESSEUR_C13_C14_C15).Text)
                '.TDetailsPhasesProduction(REDRESSEURS.R_C15).TempsAmorçageSecondes = CLng(TBTempsAmorcage.Text)
            
            'End If
                
        End With
    
        '--- analyse en fonction du redresseur pour lancer le changement de la programmation ---
        'If NumChargeEnCours = TEtatsPostes(POSTES.P_A3).NumCharge Then
        '    ChangementProgrammationRedresseur REDRESSEURS.R_C13
        'End If
        'If NumChargeEnCours = TEtatsPostes(POSTES.P_A9).NumCharge Then
        '    ChangementProgrammationRedresseur REDRESSEURS.R_A9
        'End If
        'If NumChargeEnCours = TEtatsPostes(POSTES.P_C13).NumCharge Then
        '    ChangementProgrammationRedresseur REDRESSEURS.R_C13
        'End If
        'If NumChargeEnCours = TEtatsPostes(POSTES.P_C14).NumCharge Then
        '    ChangementProgrammationRedresseur REDRESSEURS.R_C14
        'End If
        'If NumChargeEnCours = TEtatsPostes(POSTES.P_C15).NumCharge Then
        '    ChangementProgrammationRedresseur REDRESSEURS.R_C15
        'End If
    
    End If

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
        Case vbKeyDelete: EditionDetailsGammesAnodisation vbKeyBack  'simule un retour arrière (effacement)
        Case Else
    End Select
End Sub

Private Sub MSHFGDetailsGammesAnodisation_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    EditionDetailsGammesAnodisation KeyAscii  'envoi de la touche frappée
End Sub

Private Sub MSHFGDetailsGammesAnodisation_LeaveCell()
    On Error Resume Next
    MEBEditionDetailsGammesAnodisation.Visible = False
End Sub

Private Sub MSHFGDetailsGammesAnodisation_LostFocus()
    On Error Resume Next
    SFocusTableDetailsGammesAnodisation.Visible = False
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Permet l'édition des détails des gammes d'anodisation
' Entrées : KeyAscii -> Code ASCII de la touche frappée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EditionDetailsGammesAnodisation(ByRef KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim NumLigne As Integer, _
           NumColonne As Integer

    If NumChargeEnCours > 0 Then
    
        '--- affectation ---
        With MSHFGDetailsGammesAnodisation
            NumLigne = .Row
            NumColonne = .Col
        End With
    
        With TEtatsCharges(NumChargeEnCours).TGammesAnodisation.TDetailsGammesAnodisation(NumLigne)
        
            '--- pas d'édition des champs si la ligne est vide ---
            If .NumZone = 0 Then
                Exit Sub
            End If

            '--- pas d'édition des champs si pas de temps au poste ou pas d'égouttage ---
            Select Case NumColonne
            
                Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                    '--- temps au poste ---
                    If TEtatsPostes(TZones(.NumZone).NumPremierPoste).DefinitionPoste.AvecTemps = False Then
                        Exit Sub
                    End If
            
                Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                    '--- temps d'égouttage ---
                    If TEtatsPostes(TZones(.NumZone).NumPremierPoste).DefinitionPoste.AvecEgouttage = False Then
                        Exit Sub
                    End If
            
                Case Else
            End Select

            '--- édition uniquement sur les bonnes colonnes ---
            Select Case NumColonne

                Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE, _
                         COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE, _
                         COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
        
                    With MEBEditionDetailsGammesAnodisation
        
                        '--- affiche le contrôle texte au bon endroit (dans la cellule) ---
                        .Move MSHFGDetailsGammesAnodisation.Left + MSHFGDetailsGammesAnodisation.CellLeft, _
                                   MSHFGDetailsGammesAnodisation.Top + MSHFGDetailsGammesAnodisation.CellTop, _
                                   MSHFGDetailsGammesAnodisation.CellWidth, _
                                   MSHFGDetailsGammesAnodisation.CellHeight

                        '--- paramètres de contrôle texte en fonction de la cellule ---
                        .Mask = ""
                        .Text = ""
                        Select Case MSHFGDetailsGammesAnodisation.Col
                            Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE: .Mask = "##:##:##"
                            Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE: .Mask = "##:##:##"
                            Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE: .Mask = "##:##"
                            Case Else
                        End Select

                        '--- analyse du caractère qui a été tapé ---
                        Select Case KeyAscii
        
                            Case 0 To Pred(vbKeyBack), Succ(vbKeyBack) To Pred(vbKeyReturn), Succ(vbKeyReturn) To vbKeySpace
                                '--- du code 0 à l'espace (sauf retour arrière, Entrée) cela signifie une modification du texte en cours ---
                                .SelText = MSHFGDetailsGammesAnodisation.Text
                                .SelStart = 0
                                .SelLength = Len(.Text)
                                .Visible = True
                                .SetFocus
        
                            Case vbKeyBack
                                '--- touche retour arrière ---
                                .SelText = ""
                                .Visible = True
                                .SetFocus
        
                            Case vbKeyReturn
                                '--- touche Entrée ---
                                With MSHFGDetailsGammesAnodisation
                                    Select Case .Col
                                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE: .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE: .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                                            If .Row < .Rows - 1 Then .Row = .Row + 1
                                            .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
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
    
        End With
    
    End If
    
End Sub

Private Sub MEBEditionDetailsGammesAnodisation_Change()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---

    If NumChargeEnCours > 0 And InterdireEvenements = False Then

        With MSHFGDetailsGammesAnodisation

            Select Case .Col
             
                Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                    '--- temps au poste en texte ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                    '--- temps alerte en texte ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col

                Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                    '--- temps d'égouttage en texte ---
                    MemNumLigne = .Row
                    MemNumColonne = .Col
                
                Case Else
 
            End Select

        End With
    
    End If

End Sub

Private Sub MEBEditionDetailsGammesAnodisation_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    MemNumLigne = 0
    MemNumColonne = 0
    
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
                '--- flèche basse ---
                .SetFocus
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                End If
                KeyCode = 0
            
            Case vbKeyUp
                '--- flèche haute ---
                .SetFocus
                If .Row > .FixedRows Then
                    .Row = .Row - 1
                End If
                KeyCode = 0
  
            Case Else
  
        End Select
  
    End With
  
End Sub

Private Sub MEBEditionDetailsGammesAnodisation_KeyPress(KeyAscii As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    If NumChargeEnCours > 0 Then
  
        '--- affectation ---
        With MSHFGDetailsGammesAnodisation

            '--- analyse de la touche ---
            Select Case KeyAscii
    
                Case vbKeyReturn
                    '--- touche entrée ---
                    Select Case .Col
    
                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                            '--- temps au poste en texte ---
                            .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                        
                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                            '--- temps alerte en texte ---
                            .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                        
                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                            '--- temps d'égouttage en texte ---
                            If .Row < .Rows - 1 Then .Row = .Row + 1
                            .Col = COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                            
                        Case Else
                    End Select
    
                    '--- mettre le focus sur le tableau ---
                    .SetFocus
                    KeyAscii = 0

                Case Else
                    Select Case .Col
                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 8
                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 8
                        Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE: FiltreToucheASCII KeyAscii, DONNEES.D_NBR_NATURELS, 5
                        Case Else
                    End Select
    
            End Select
    
        End With

    End If

End Sub

Private Sub MEBEditionDetailsGammesAnodisation_LostFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim TexteComplet As String, _
            TexteSansMasque As String
    Dim TempsAuPosteTexteProvisoire As String      'temps au poste texte provisoire pour la modification avec
                                                                                   'confirmation des gammes d'anodisation en cours de traitement
    Dim TempsAlerteTexteProvisoire As String          'temps d'alerte texte provisoire pour la modification avec
                                                                                   'confirmation des gammes d'anodisation en cours de traitement
    Dim TempsEgouttageTexteProvisoire As String   'temps d'égouttage texte provisoire pour la modification avec
                                                                                   'confirmation des gammes d'anodisation en cours de traitement
    
    '--- demande de validation des changements ---
    If NumChargeEnCours > 0 And MemNumLigne > 0 And MemNumColonne > 0 Then
        
        '--- affectation ---
        With MEBEditionDetailsGammesAnodisation
            TexteComplet = .Text
            TexteSansMasque = .ClipText
        End With
        
        Select Case MemNumColonne
        
            Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_AU_POSTE_TEXTE
                '--- temps au poste en texte ---
                TempsAuPosteTexteProvisoire = Replace(TexteComplet, "_", "0")
                
                If AppelFenetre(F_MESSAGE, _
                                        TITRE_MESSAGES, _
                                        vbCrLf & "cs|MODIFICATION DU TEMPS AU POSTE" & vbCrLf & vbCrLf & _
                                        "Zone = " & TGammesAnodisation.TDetailsGammesAnodisation(MemNumLigne).LibelleZone & vbCrLf & _
                                        "Temps demandé = " & TempsAuPosteTexteProvisoire & vbCrLf & vbCrLf & _
                                        "cs|Voulez-vous réellement valider ce temps ?", _
                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                        TYPES_BOUTONS.T_OUI_NON, _
                                        EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                    '--- changer le temps au poste directement dans les états des charges ---
                    With TEtatsCharges(NumChargeEnCours).TGammesAnodisation.TDetailsGammesAnodisation(MemNumLigne)
                        .TempsAuPosteTexte = TempsAuPosteTexteProvisoire
                        .TempsAuPosteSecondes = CTempsTexteEnSecondes(.TempsAuPosteTexte)
                        .FinDuTempsPosteReel = False
                    End With
                
                Else
                
                    '--- replacer le focus sur la grille au bon endroit ---
                    With MSHFGDetailsGammesAnodisation
                        .Row = MemNumLigne
                        .Col = MemNumColonne
                        .SetFocus
                    End With
                
                End If
            
            Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_ALERTE_TEXTE
                '--- temps d'alerte en texte ---
                TempsAlerteTexteProvisoire = Replace(TexteComplet, "_", "0")
                
                If AppelFenetre(F_MESSAGE, _
                                        TITRE_MESSAGES, _
                                        vbCrLf & "cs|MODIFICATION DU TEMPS D'ALERTE" & vbCrLf & vbCrLf & _
                                        "Zone = " & TGammesAnodisation.TDetailsGammesAnodisation(MemNumLigne).LibelleZone & vbCrLf & _
                                        "Temps demandé = " & TempsAlerteTexteProvisoire & vbCrLf & vbCrLf & _
                                        "cs|Voulez-vous réellement valider ce temps ?", _
                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                        TYPES_BOUTONS.T_OUI_NON, _
                                        EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                    '--- changer le temps au poste directement dans les états des charges ---
                    With TEtatsCharges(NumChargeEnCours).TGammesAnodisation.TDetailsGammesAnodisation(MemNumLigne)
                        .TempsAlerteTexte = TempsAlerteTexteProvisoire
                        .TempsAlerteSecondes = CTempsTexteEnSecondes(.TempsAlerteTexte)
                        .DebutAlertePosteReel = False
                    End With
                
                Else
                
                    '--- replacer le focus sur la grille au bon endroit ---
                    With MSHFGDetailsGammesAnodisation
                        .Row = MemNumLigne
                        .Col = MemNumColonne
                        .SetFocus
                    End With
                
                End If
        
            Case COLONNES_DETAILS_GAMMES_ANODISATION.C_TEMPS_EGOUTTAGE_TEXTE
                '--- temps d'égouttage en texte ---
                TempsEgouttageTexteProvisoire = Replace(TexteComplet, "_", "0")
                
                If AppelFenetre(F_MESSAGE, _
                                        TITRE_MESSAGES, _
                                        vbCrLf & "cs|MODIFICATION DU TEMPS D'EGOUTTAGE" & vbCrLf & vbCrLf & _
                                        "Zone = " & TGammesAnodisation.TDetailsGammesAnodisation(MemNumLigne).LibelleZone & vbCrLf & _
                                        "Temps demandé = " & TempsEgouttageTexteProvisoire & vbCrLf & vbCrLf & _
                                        "cs|Voulez-vous réellement valider ce temps ?", _
                                        TYPES_MESSAGES.T_AVERTISSEMENT, _
                                        TYPES_BOUTONS.T_OUI_NON, _
                                        EMPLACEMENT_FOCUS.E_SUR_NON) = vbYes Then
                
                    '--- changer le temps au poste directement dans les états des charges ---
                    With TEtatsCharges(NumChargeEnCours).TGammesAnodisation.TDetailsGammesAnodisation(MemNumLigne)
                        .TempsEgouttageTexte = TempsEgouttageTexteProvisoire
                        .TempsEgouttageSecondes = CTempsTexteEnSecondes(.TempsEgouttageTexte)
                    End With
                
                Else
                
                    '--- replacer le focus sur la grille au bon endroit ---
                    With MSHFGDetailsGammesAnodisation
                        .Row = MemNumLigne
                        .Col = MemNumColonne
                        .SetFocus
                    End With
                
                End If
                
            Case Else
    
        End Select

    End If
    
    '--- focus ---
    SFocusTableDetailsGammesAnodisation.Visible = False
        
    '--- rendre le contrôle texte invisible ---
    MEBEditionDetailsGammesAnodisation.Visible = False

    '--- construction de la grille ---
    GestionGammesAnodisation GG_TRANSFERT_DONNEES
    GestionGammesAnodisation GG_AFFICHAGE
    
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
            MemColonne As Integer, _
            PtrLigne As Integer, _
            DerniereLigneAfficher As Integer
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
                    .RowHeight(a) = 750                    'épaisseur des lignes
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
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                PtrLigne = 1
                For a = LBound(TDetailsFichesProduction()) To UBound(TDetailsFichesProduction())
                    
                    With TEtatsCharges(NumChargeEnCours).TDetailsFichesProduction(a)
                        
                        If .NumPoste >= POSTES.P_CHGT_1 And .NumPoste <= DERNIER_POSTE Then
                        
                            '--- numéro et nom du poste ---
                            TDetailsFichesProduction(PtrLigne).NomPoste = TEtatsPostes(.NumPoste).DefinitionPoste.NomPoste
                            TDetailsFichesProduction(PtrLigne).NumPoste = .NumPoste
                            
                            '--- temps réel au poste ---
                            TDetailsFichesProduction(PtrLigne).TempsReelPoste = "Entrée le " & Format(.DateEntreePoste, FORMAT_DATE_HEURE_1) & vbCr
                            If .DateSortiePoste = Empty Then
                                TDetailsFichesProduction(PtrLigne).TempsReelPoste = TDetailsFichesProduction(PtrLigne).TempsReelPoste & "-" & vbCr & "-"
                            Else
                                TempsEnSecondes = DateDiff("s", .DateEntreePoste, .DateSortiePoste)
                                TDetailsFichesProduction(PtrLigne).TempsReelPoste = TDetailsFichesProduction(PtrLigne).TempsReelPoste & _
                                                                                                                          "Sortie le  " & Format(.DateSortiePoste, FORMAT_DATE_HEURE_1) & vbCr & _
                                                                                                                          "Temps = " & CTemps2(TempsEnSecondes)
                            End If
                            
                            '--- temps réel d'égouttage ---
                            If .DateDebutEgouttage = Empty Then
                                TDetailsFichesProduction(PtrLigne).TempsReelEgouttage = "-" & vbCr
                            Else
                                TDetailsFichesProduction(PtrLigne).TempsReelEgouttage = "Début le " & Format(.DateDebutEgouttage, FORMAT_DATE_HEURE_1) & vbCr
                            End If
                            If .DateFinEgouttage = Empty Then
                                TDetailsFichesProduction(PtrLigne).TempsReelEgouttage = TDetailsFichesProduction(PtrLigne).TempsReelEgouttage & "-" & vbCr & "-"
                            Else
                                TempsEnSecondes = DateDiff("s", .DateDebutEgouttage, .DateFinEgouttage)
                                TDetailsFichesProduction(PtrLigne).TempsReelEgouttage = TDetailsFichesProduction(PtrLigne).TempsReelEgouttage & _
                                                                                                                                 "Fin le  " & Format(.DateFinEgouttage, FORMAT_DATE_HEURE_1) & vbCr & _
                                                                                                                                 "Temps = " & CTemps2(TempsEnSecondes)
                            End If
                            
                            '--- températures ---
                            If .TemperatureEnEntree = 0 Then
                                TDetailsFichesProduction(PtrLigne).Temperatures = "-" & vbCr & "-"
                            Else
                                TDetailsFichesProduction(PtrLigne).Temperatures = "En entrant : " & Format(.TemperatureEnEntree, FORMAT_TEMPERATURE_1_DECIMALE_UNITE)
                                If .TemperatureEnSortie = 0 Then
                                    TDetailsFichesProduction(PtrLigne).Temperatures = TDetailsFichesProduction(PtrLigne).Temperatures & vbCr & "-"
                                Else
                                    TDetailsFichesProduction(PtrLigne).Temperatures = TDetailsFichesProduction(PtrLigne).Temperatures & vbCr & _
                                                                                                                       "En sortant : " & Format(.TemperatureEnSortie, FORMAT_TEMPERATURE_1_DECIMALE_UNITE)
                                End If
                            End If
                            
                            '--- redresseur ---
                            If .URedresseur = 0 Then
                                TDetailsFichesProduction(PtrLigne).Redresseur = "-" & vbCr & "-"
                            Else
                                Select Case .NumPoste
                                    'Case POSTES.P_A3
                                        'If .SensRedresseur = SENS_REDRESSEUR.R_EN_CATHODIQUE_OU_POLARISATION Then
                                        '    TDetailsFichesProduction(PtrLigne).Redresseur = TEXTE_CATHODIQUE
                                        'Else
                                        '   TDetailsFichesProduction(PtrLigne).Redresseur = TEXTE_ANODIQUE
                                        'End If
                                    'Case POSTES.P_A9
                                        'TDetailsFichesProduction(PtrLigne).Redresseur = TEXTE_CATHODIQUE
                                    Case POSTES.P_C13, POSTES.P_C14, POSTES.P_C15, POSTES.P_C16
                                        'If .SensRedresseur = SENS_REDRESSEUR.R_EN_CATHODIQUE_OU_POLARISATION Then
                                        '    TDetailsFichesProduction(PtrLigne).Redresseur = TEXTE_POLARISATION
                                        'Else
                                         '   TDetailsFichesProduction(PtrLigne).Redresseur = TEXTE_AMORCAGE
                                        'End If
                                    Case Else
                                End Select
                                TDetailsFichesProduction(PtrLigne).Redresseur = TDetailsFichesProduction(PtrLigne).Redresseur & vbCr & "U = " & Format(.URedresseur, FORMAT_TENSION_1_DECIMALE_UNITE)
                                If .IRedresseur = 0 Then
                                    TDetailsFichesProduction(PtrLigne).Redresseur = TDetailsFichesProduction(PtrLigne).Redresseur & vbCr & "-"
                                Else
                                    TDetailsFichesProduction(PtrLigne).Redresseur = TDetailsFichesProduction(PtrLigne).Redresseur & vbCr & _
                                                                                                                    "I = " & Format(.IRedresseur, FORMAT_INTENSITE_ENTIER_UNITE)
                                End If
                            End If
                            
                            '--- analyseur ---
                            If .AnalyseurEnEntree = 0 Then
                                TDetailsFichesProduction(PtrLigne).Analyseur = "-" & vbCr & "-"
                            Else
                                TDetailsFichesProduction(PtrLigne).Analyseur = "En entrant : " & Format(.AnalyseurEnEntree, FORMAT_ANALYSEUR_UNITE)
                                If .AnalyseurEnSortie = 0 Then
                                    TDetailsFichesProduction(PtrLigne).Analyseur = TDetailsFichesProduction(PtrLigne).Analyseur & vbCr & "-"
                                Else
                                    TDetailsFichesProduction(PtrLigne).Analyseur = TDetailsFichesProduction(PtrLigne).Analyseur & vbCr & _
                                                                                                                   "En sortant : " & Format(.AnalyseurEnSortie, FORMAT_ANALYSEUR_UNITE)
                                End If
                            End If
                            
                            '--- alarmes de poste ---
                            TDetailsFichesProduction(PtrLigne).AlarmesPoste = DecodeAlarmesPoste(.AlarmesPoste)
                            
                            Inc PtrLigne
                        
                        Else
                        
                            Exit For
                        
                        End If
                    
                    End With
                
                Next a
            End If

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
                        
                        '--- affectation de la dernière ligne afficher ---
                        DerniereLigneAfficher = a
                        
                        '--- affichage dans la grille ---
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

                '--- rendre toujours visible indication du pointeur en cours ---
                If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                    If MemNumChargeEnCoursPourAffichage <> NumChargeEnCours Then
                        If .RowIsVisible(DerniereLigneAfficher) = False Then
                            If DerniereLigneAfficher <= NBR_LIGNES_DETAILS_GAMMES_PRODUCTION Then
                                .TopRow = DerniereLigneAfficher
                            End If
                        End If
                        MemNumChargeEnCoursPourAffichage = NumChargeEnCours 'mémoire anti-rebond
                    End If
                End If
                
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
' Rôle      : Affichage de la globalité des temps
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageGlobaliteTemps()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim TempsAvantPostePrincipalSansPontsSecondes As Long        'temps avant le poste principal sans les ponts en secondes
    Dim TempsPostePrincipalSansPontsSecondes As Long                 'temps au poste principal sans les ponts en secondes
    Dim TempsApresPostePrincipalSansPontsSecondes As Long        'temps aprés le poste principal sans les ponts en secondes
    Dim TempsTotalPostesSansPontsSecondes As Long                      'temps total des postes sans les ponts en secondes
    Dim TempsTotalEgouttagesSansPontsSecondes As Long               'temps total des égouttages sans les ponts en secondes
    Dim TempsTotalGammeSansPontsSecondes As Long                     'temps total de la gamme sans les ponts en secondes

    Dim TempsMouvementsAvantPostePrincipalSecondes As Long     'temps des mouvements avant le poste principal en secondes
    Dim TempsAvantPostePrincipalAvecPontsSecondes As Long         'temps avant le poste principal avec les ponts en secondes
    Dim TempsPostePrincipalAvecPontsSecondes As Long                  'temps au poste principal avec les ponts en secondes
    Dim TempsMouvementsApresPostePrincipalSecondes As Long     'temps des mouvements après le poste principal en secondes
    Dim TempsApresPostePrincipalAvecPontsSecondes As Long         'temps après le poste principal avec les ponts en secondes
    Dim TempsTotalPostesAvecPontsSecondes As Long                       'temps total des postes avec les ponts en secondes
    Dim TempsTotalEgouttagesAvecPontsSecondes As Long                'temps total des égouttages avec les ponts en secondes
    Dim TempsTotalMouvementsSecondes As Long                              'temps total des mouvements en secondes
    Dim TempsTotalGammeAvecPontsSecondes As Long                      'temps total de la gamme avec les ponts en secondes

    Dim TempsAvantPostePrincipalSansPontsTexte As String              'temps avant le poste principal sans les ponts en texte au format HH:MM:SS
    Dim TempsPostePrincipalSansPontsTexte As String                       'temps au poste principal sans les ponts en texte au format HH:MM:SS
    Dim TempsApresPostePrincipalSansPontsTexte As String              'temps aprés poste principal sans les ponts en texte au format HH:MM:SS
    Dim TempsTotalPostesSansPontsTexte As String                            'temps total des postes sans les ponts en texte au format HH:MM:SS
    Dim TempsTotalEgouttagesSansPontsTexte As String                     'temps total des égouttages sans les ponts en texte au format HH:MM:SS
    Dim TempsTotalGammeSansPontsTexte As String                           'temps total de la gamme sans les ponts en texte au format HH:MM:SS
        
    Dim TempsMouvementsAvantPostePrincipalTexte As String           'temps des mouvements avant le poste principal au format HH:MM:SS
    Dim TempsAvantPostePrincipalAvecPontsTexte As String               'temps avant le poste principal avec les ponts au format HH:MM:SS
    Dim TempsAnodisationAvecPontsTexte As String                             'temps au poste principal avec les ponts au format HH:MM:SS
    Dim TempsMouvementsApresPostePrincipalTexte As String           'temps des mouvements après le poste principal au format HH:MM:SS
    Dim TempsApresPostePrincipalAvecPontsTexte As String               'temps après le poste principal avec les ponts au format HH:MM:SS
    Dim TempsTotalPostesAvecPontsTexte As String                             'temps total des postes avec les ponts au format HH:MM:SS
    Dim TempsTotalEgouttagesAvecPontsTexte As String                      'temps total des égouttages avec les ponts au format HH:MM:SS
    Dim TempsTotalMouvementsTexte As String                                    'temps total des mouvements au format HH:MM:SS
    Dim TempsTotalGammeAvecPontsTexte As String                            'temps total de la gamme avec les ponts au format HH:MM:SS
    
    
    '********************************************************************************************************************
    '                                                                   EXTRACTION DES TEMPS
    '********************************************************************************************************************
    '--- calcul des temps de la gamme d'anodisation sans les ponts à condition de disposer d'un n° de charge en cours ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
        With TEtatsCharges(NumChargeEnCours)
            
            '--- calcul des temps SANS les ponts ---
            CalculTempsGammeAnodisationSansPonts .TGammesAnodisation, _
                                                                           TempsAvantPostePrincipalSansPontsSecondes, _
                                                                           TempsPostePrincipalSansPontsSecondes, _
                                                                           TempsApresPostePrincipalSansPontsSecondes, _
                                                                           TempsTotalPostesSansPontsSecondes, _
                                                                           TempsTotalEgouttagesSansPontsSecondes, _
                                                                           TempsTotalGammeSansPontsSecondes
            
            '--- calcul des temps AVEC les ponts ---
            CalculTempsGammeAnodisationAvecPonts .TGammesAnodisation, _
                                                                         TempsMouvementsAvantPostePrincipalSecondes, _
                                                                         TempsAvantPostePrincipalAvecPontsSecondes, _
                                                                         TempsPostePrincipalAvecPontsSecondes, _
                                                                         TempsMouvementsApresPostePrincipalSecondes, _
                                                                         TempsApresPostePrincipalAvecPontsSecondes, _
                                                                         TempsTotalPostesAvecPontsSecondes, _
                                                                         TempsTotalEgouttagesAvecPontsSecondes, _
                                                                         TempsTotalMouvementsSecondes, _
                                                                         TempsTotalGammeAvecPontsSecondes
        
        End With
    End If
    
    '********************************************************************************************************************
    '                                             AFFICHAGE DES TEMPS SANS LES PONTS (GAMME)
    '********************************************************************************************************************
    
    '--- affichage du temps avant Anodisation sans les ponts ---
    If TempsAvantPostePrincipalSansPontsSecondes = 0 Then
        TempsAvantPostePrincipalSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsAvantPostePrincipalSansPontsTexte = CTemps2(TempsAvantPostePrincipalSansPontsSecondes)
    End If
    AffichageTexte LTempsAvantPostePrincipalSansPonts, TempsAvantPostePrincipalSansPontsTexte
    
    '--- affichage du temps au Anodisation sans les ponts (identique avec les ponts) ---
    If TempsPostePrincipalSansPontsSecondes = 0 Then
        TempsPostePrincipalSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsPostePrincipalSansPontsTexte = CTemps2(TempsPostePrincipalSansPontsSecondes)
    End If
    AffichageTexte LTempsPostePrincipalSansPonts, TempsPostePrincipalSansPontsTexte

    '--- affichage du temps après Anodisation sans les ponts ---
    If TempsApresPostePrincipalSansPontsSecondes = 0 Then
        TempsApresPostePrincipalSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsApresPostePrincipalSansPontsTexte = CTemps2(TempsApresPostePrincipalSansPontsSecondes)
    End If
    AffichageTexte LTempsApresPostePrincipalSansPonts, TempsApresPostePrincipalSansPontsTexte

    '--- affectation du temps total des postes sans les ponts en texte ---
    If TempsTotalPostesSansPontsSecondes = 0 Then
        TempsTotalPostesSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalPostesSansPontsTexte = CTemps2(TempsTotalPostesSansPontsSecondes)
    End If
    
    '--- affectation du temps total des égouttages sans les ponts en texte ---
    If TempsTotalEgouttagesSansPontsSecondes = 0 Then
        TempsTotalEgouttagesSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalEgouttagesSansPontsTexte = CTemps2(TempsTotalEgouttagesSansPontsSecondes)
    End If
    
    '--- affichage du temps total de la gamme sans les ponts ---
    If TempsTotalGammeSansPontsSecondes = 0 Then
        TempsTotalGammeSansPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalGammeSansPontsTexte = CTemps2(TempsTotalGammeSansPontsSecondes)
    End If
    AffichageTexte LTempsTotalGammeSansPonts, TempsTotalGammeSansPontsTexte
    
    '********************************************************************************************************************
    '                                  AFFICHAGE DES TEMPS AVEC LES PONTS (PAR APPRENTISSAGE)
    '********************************************************************************************************************
    
    '--- affichage du temps des mouvements avant Anodisation ---
    If TempsMouvementsAvantPostePrincipalSecondes = 0 Then
        TempsMouvementsAvantPostePrincipalTexte = PAS_DE_TEMPS
    Else
        TempsMouvementsAvantPostePrincipalTexte = CTemps2(TempsMouvementsAvantPostePrincipalSecondes)
    End If
    AffichageTexte LTempsMouvementsAvantPostePrincipal, TempsMouvementsAvantPostePrincipalTexte
    
    '--- affichage du temps avant Anodisation avec les ponts ---
    If TempsAvantPostePrincipalAvecPontsSecondes = 0 Then
        TempsAvantPostePrincipalAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsAvantPostePrincipalAvecPontsTexte = CTemps2(TempsAvantPostePrincipalAvecPontsSecondes)
    End If
    AffichageTexte LTempsAvantPostePrincipalAvecPonts, TempsAvantPostePrincipalAvecPontsTexte
    
    '--- affichage du temps au Anodisation (identique aux valeurs définies dans la gamme) ---
    If TempsPostePrincipalAvecPontsSecondes = 0 Then
        TempsAnodisationAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsAnodisationAvecPontsTexte = CTemps2(TempsPostePrincipalAvecPontsSecondes)
    End If
    AffichageTexte LTempsPostePrincipalAvecPonts, TempsAnodisationAvecPontsTexte
    
    '--- affichage du temps des mouvements après Anodisation ---
    If TempsMouvementsApresPostePrincipalSecondes = 0 Then
        TempsMouvementsApresPostePrincipalTexte = PAS_DE_TEMPS
    Else
        TempsMouvementsApresPostePrincipalTexte = CTemps2(TempsMouvementsApresPostePrincipalSecondes)
    End If
    AffichageTexte LTempsMouvementsApresPostePrincipal, TempsMouvementsApresPostePrincipalTexte
    
    '--- affichage du temps après Anodisation ---
    If TempsApresPostePrincipalAvecPontsSecondes = 0 Then
        TempsApresPostePrincipalAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsApresPostePrincipalAvecPontsTexte = CTemps2(TempsApresPostePrincipalAvecPontsSecondes)
    End If
    AffichageTexte LTempsApresPostePrincipalAvecPonts, TempsApresPostePrincipalAvecPontsTexte
    
    '--- affichage du temps total de la gamme ---
    If TempsTotalGammeAvecPontsSecondes = 0 Then
        TempsTotalGammeAvecPontsTexte = PAS_DE_TEMPS
    Else
        TempsTotalGammeAvecPontsTexte = CTemps2(TempsTotalGammeAvecPontsSecondes)
    End If
    AffichageTexte LTempsTotalGammeAvecPonts, TempsTotalGammeAvecPontsTexte
    
    '--- affichage du temps total des mouvements ---
    If TempsTotalMouvementsSecondes = 0 Then
        TempsTotalMouvementsTexte = PAS_DE_TEMPS
    Else
        TempsTotalMouvementsTexte = CTemps2(TempsTotalMouvementsSecondes)
    End If
    AffichageTexte LTempsTotalMouvements, TempsTotalMouvementsTexte
    
    '********************************************************************************************************************
    '                                             AFFICHAGE DES PREVISIONS EN TEMPS REEL
    '********************************************************************************************************************
    
    'LPrevisionsTempsReel(0).Caption = "Heure de prise au chargement vers "
    'LPrevisionsTempsReel(2).Caption = "Dépose au poste principal (anodisation, satinage,...) vers "
    'LPrevisionsTempsReel(3).Caption = "Sortie du poste principal (anodisation, satinage,...) vers "
    'LPrevisionsTempsReel(5).Caption = "Sortie de la ligne vers "
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Visualisation des différents états de la charge gérée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EtatsChargeGeree()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Static MemNumChargeEnCours As Integer
    
    '--- affichage dans les tableaux ---
    If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
        
        '--- lancer l'affichage une seule fois ---
        If MemNumChargeEnCours <> NumChargeEnCours Then
            AffichageChargeGeree
            GestionDetailsCharges GG_TRANSFERT_DONNEES
            GestionDetailsCharges GG_AFFICHAGE
            MemNumChargeEnCours = NumChargeEnCours
        End If
        
        '--- rafraichir régulièrement ---
        GestionEtatsCharges GG_TRANSFERT_DONNEES
        GestionEtatsCharges GG_AFFICHAGE
        GestionGammesAnodisation GG_TRANSFERT_DONNEES
        GestionGammesAnodisation GG_AFFICHAGE
        GestionDetailsPhasesProduction GG_TRANSFERT_DONNEES
        GestionDetailsPhasesProduction GG_AFFICHAGE
        GestionDetailsFichesProduction GG_TRANSFERT_DONNEES
        GestionDetailsFichesProduction GG_AFFICHAGE
        AffichageDateFinDansLePoste
        AffichageGlobaliteTemps
        
    End If
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Gestion des détails des phases de la production
' Entrées : EtatSouhaite -> Fonction de l'énumération GESTION_GRILLES
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionDetailsPhasesProduction(ByVal EtatSouhaite As GESTION_GRILLES)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---

    '--- déclaration ---
    Dim AffichageZoneRedresseur As Boolean             'indique qu'il faut ou non afficher la zone redresseur
    Dim a As Integer                                                       'pour les boucles FOR...NEXT
    Dim NumZone As Integer                                         'n° de la zone
    Dim ModeUouI As Integer                                         'pour le passage par référence
    
    Dim Texte As String

    Select Case EtatSouhaite

        Case GESTION_GRILLES.GG_INITIALISATION, _
                 GESTION_GRILLES.GG_VIDAGE
            '--- initialisation du tableau ---
            Erase TDetailsPhasesProduction()
        
            '--- initialisation des champs redresseurs ---
            InitialisationChampsRedresseur
            
            '--- effacement de la zone des redresseurs ---
            FRedresseurs.Visible = False
        
        Case GESTION_GRILLES.GG_TRANSFERT_DONNEES
            '--- transfert des données dans le tableau ---
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
                
                With TEtatsCharges(NumChargeEnCours)
                    
                    For a = LBound(.TDetailsPhasesProduction()) To UBound(.TDetailsPhasesProduction())
                        
                        '--- mode tension ou intensité ---
                        TDetailsPhasesProduction(a).ModeUouI = .ModeUouI
                        
                        '--- temps des phases, tensions et intensités ---
                        With .TDetailsPhasesProduction(a)
                            TDetailsPhasesProduction(a).TempsPhase = .TempsPhase
                            TDetailsPhasesProduction(a).UPhase = .UPhase
                            TDetailsPhasesProduction(a).IPhase = .IPhase
                        End With
                    
                    Next a
                
                End With
            
            End If
            
        Case GESTION_GRILLES.GG_AFFICHAGE
            '--- affichage des valeurs des redresseurs ---
            If NumChargeEnCours >= CHARGES.C_NUM_MINI And NumChargeEnCours <= CHARGES.C_NUM_MAXI Then
            
                '--- lancer le contrôle une seule fois ---
                If MemNumChargeEnCoursPourPhasesProduction <> NumChargeEnCours Then
        
                    '--- ne pas afficher la partie redresseur par défaut ---
                    'AffichageZoneRedresseur = False
                                       
                    With TEtatsCharges(NumChargeEnCours)
                        
                        '--- interdire les évènements ---
                        InterdireEvenements = True
                        
                        '--- rendre visible le redresseur se trouvant dans la gamme ---
                        For a = LBound(.TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(.TGammesAnodisation.TDetailsGammesAnodisation())
                            
                            '--- affectation du numéro de zone ---
                            NumZone = .TGammesAnodisation.TDetailsGammesAnodisation(a).NumZone
                            ' SZP2024
                            If NumZone >= LIMITE_BASSE_ZONES And NumZone <= LIMITE_HAUTE_ZONES Then
                            
                                '--- affichage de la partie redresseur ---
                                Call Log("TZones(NumZone).Codezone =" & TZones(NumZone).Codezone & " !!!!!!!!!!")
                                If TZones(NumZone).NumZone = NUMZONE_ANO Then
                                    AffichageZoneRedresseur = True
                                End If
                            
                            End If
                        
                        Next a
            
                        '--- affichage des valeurs de programmation pour le redresseur ---
                        If AffichageZoneRedresseur = True Then
                            
                            '--- mode U ou I en mode tension ---
                            ModeUouI = .ModeUouI
                            Call LModeUouI_Click(ModeUouI)
                            
                            '--- affichage des temps, tensions, intensités des phases ---
                            For a = LBound(TDetailsPhasesProduction()) To UBound(TDetailsPhasesProduction())
                                With TDetailsPhasesProduction(a)
                                    MEBTempsPhases(a).Text = Right(CTemps2(.TempsPhase), 7)
                                    TBTensionsPhases(a).Text = Format(.UPhase, FORMAT_TENSION_1_DECIMALE)
                                    TBIntensitesPhases(a).Text = Format(.IPhase, FORMAT_INTENSITE_ENTIER)
                                End With
                            Next a
                        
                        End If
                        
                        '--- calcul du temps total de la gamme redresseur ---
                        LTempsTotalGammeRedresseur.Caption = Right(CTemps2(CalculTempsTotalGammeRedresseur()), 7)
                        
                        '--- autoriser les évènements ---
                        InterdireEvenements = True
                
                    End With
                    
                    '--- affichage de la zone redresseur ---
                    FRedresseurs.Visible = AffichageZoneRedresseur
                    FRedresseurs.Refresh
                    
                    '--- affectation ---
                    MemNumChargeEnCoursPourPhasesProduction = NumChargeEnCours
                
                End If
        
            End If

        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Modification des dimensions du cadre de la gamme
' Entrées :
' Sorties  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ModificationDimensionsCadreGamme()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const HAUTEUR_BASE_CADRE As Long = 3495
    Const HAUTEUR_ETENDU_CADRE As Long = 7935

    '--- déclaration ---
    
    '--- modification des dimensions ---
    With FGammeAnodisation
        If .Height = HAUTEUR_BASE_CADRE Then
            .Height = HAUTEUR_ETENDU_CADRE
            LNouveauPointeur(0).Visible = True
            LNouveauPointeur(1).Visible = True
            SNouveauPointeur.Visible = True
            TBNouveauPointeur.SetFocus
        Else
            .Height = HAUTEUR_BASE_CADRE
            LNouveauPointeur(0).Visible = False
            LNouveauPointeur(1).Visible = False
            SNouveauPointeur.Visible = False
            MSHFGDetailsGammesAnodisation.SetFocus
        End If
    End With

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
' Rôle      : Gére l'image tampon (affichage de l'image tampon à l'écran)
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub GestionImageTampon()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim RDestination As RECT  'coordonnées du rectangle de destination
    
    '--- récupération des coordonnées écran de l'image de la ligne ---
    Call ObjDX.GetWindowRect(PBImageLigne.hwnd, RDestination)

    '--- transfert de l'image tampon à l'écran ---
    Call ObjDDSEcran.Blt(RDestination, ObjDDSImageTampon, RImageTampon, DDBLT_WAIT)
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation de DirectX
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationDirectX()
        
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
            
    '--- dimensionnement de l'image de la ligne ---
    With PBImageLigne
        .Width = LONGUEUR_IMAGE_LIGNE
        .Height = HAUTEUR_IMAGE_LIGNE
    End With
    
    '--- création de l'objet direct draw ---
    Set ObjDD = ObjDX.DirectDrawCreate("")
    
    '--- niveau de coopération avec l'écran ---
    Call ObjDD.SetCooperativeLevel(Me.hwnd, DDSCL_NORMAL)
    
    '--- description de la surface de l'écran ---
    With DDSDEcran
        .lFlags = DDSD_CAPS
        .ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    End With
    
    '--- création de la surface ---
    Set ObjDDSEcran = ObjDD.CreateSurface(DDSDEcran)
    
    '--- création de l'objet clipper pour utiliser que certaines régions de l'écran ---
    Set ObjDDClip = ObjDD.CreateClipper(0)
    
    '--- association de l'image à l'objet clipper ---
    ObjDDClip.SetHWnd PBImageLigne.hwnd
    
    '--- attachement du clipping à l'écran ---
    ObjDDSEcran.SetClipper ObjDDClip
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation des surfaces
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationSurfaces()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- description de l'image de la ligne ---
    With DDSDImageLigne
        .lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
        .ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
        .lWidth = PBImageLigne.Width
        .lHeight = PBImageLigne.Height
    End With
    
    '--- création de la surface et chargement de l'image de la ligne ---
    Set ObjDDSImageLigne = ObjDD.CreateSurfaceFromFile(RepImagesAnodisation & "Synoptique.bmp", DDSDImageLigne)
    
    '--- coordonnées du rectangle du synoptique ---
    With RImageLigne
        .Left = 0
        .Top = 0
        .Right = DDSDImageLigne.lWidth
        .Bottom = DDSDImageLigne.lHeight
    End With
                                                                        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche la totalité des données des redresseurs
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageDonneesRedresseurs()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                              'pour les boucles FOR...NEXT
    
    Dim CouleurFond As Long
    Dim CouleurPlan As Long
    
    Dim Texte As String                          'repésente un texte quelconque

    For a = REDRESSEURS.R_C13 To REDRESSEURS.R_C16
    
        With TEtatsRedresseurs(a)
        
            '***********************************************************************************************************************************************
            '                                                                               SUR LE DESSIN DU REDRESSEUR
            '***********************************************************************************************************************************************
            
            '--- mode du redresseur ---
            Select Case .ModeRedresseur
                
                Case MODES_REDRESSEUR.MR_MANUEL
                    Texte = "MANUEL"
                    CouleurFond = COULEURS.JAUNE_1: CouleurPlan = COULEURS.ROUGE_5
                
                Case MODES_REDRESSEUR.MR_AUTOMATIQUE
                    Texte = "AUTO."
                    CouleurFond = COULEURS.JAUNE_1: CouleurPlan = COULEURS.ROUGE_5
                
                Case Else
                    Texte = "-"
                    CouleurFond = COULEURS.JAUNE_1: CouleurPlan = COULEURS.ROUGE_5
            
            End Select
            AffichageTexte LModeRedresseurs(a), Texte, CouleurFond, CouleurPlan
            
            '--- tension ---
            If .EtatRedresseur = ER_ARRET Then
                AffichageTexte LURedresseurs(a), "-"
            Else
                AffichageTexte LURedresseurs(a), Format(.U, FORMAT_TENSION_1_DECIMALE_UNITE)
            End If
            
            '--- intensité ---
            If .EtatRedresseur = ER_ARRET Then
                AffichageTexte LIRedresseurS(a), "-"
            Else
                AffichageTexte LIRedresseurS(a), Format(.I, FORMAT_INTENSITE_ENTIER_UNITE)
            End If
            
            '--- temps restant du cycle ---
            If .TempsRestantCycle = 0 Then
                Texte = "-"
            Else
                Texte = CTemps(.TempsRestantCycle)
            End If
            AffichageTexte LTempsRestantCycle(a), Texte
            
            '--- phase en cours ---
             If .NumPhaseEnCours = 0 Then
                Texte = "-"
            Else
                Texte = .NumPhaseEnCours
            End If
            
            '--- couleurs de la phase en cours pour signaler un défaut ---
            If .TEntreesAPI.M_DefautGeneral = True Or _
                .TEntreesAPI.M_DelaiTropLongMarcheRedresseur = True Or _
                .TEntreesAPI.M_IntensiteInstable = True Or _
                .TEntreesAPI.M_IntensiteNonAtteinte = True Then
                AffichageTexte LNumPhaseEnCours(a), Texte, COULEURS.ROUGE_3, COULEURS.JAUNE_3
            Else
                AffichageTexte LNumPhaseEnCours(a), Texte, COULEURS.VERT_3, COULEURS.NOIR
            End If
            
        End With

    Next a

End Sub



VERSION 5.00
Begin VB.Form FTempsMouvements 
   ClientHeight    =   15090
   ClientLeft      =   1950
   ClientTop       =   825
   ClientWidth     =   16080
   Icon            =   "FTempsMouvements.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   15090
   ScaleWidth      =   16080
   WindowState     =   2  'Maximized
   Begin VB.PictureBox PBDeplacementFenetre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   13515
      Index           =   0
      Left            =   0
      ScaleHeight     =   13515
      ScaleWidth      =   16080
      TabIndex        =   3
      Top             =   375
      Width           =   16080
      Begin VB.PictureBox PBDeplacementFenetre 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   12975
         Index           =   1
         Left            =   0
         ScaleHeight     =   12975
         ScaleWidth      =   28635
         TabIndex        =   4
         Top             =   0
         Width           =   28635
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
            Index           =   2
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   88
            Top             =   7800
            Width           =   4455
         End
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
            Index           =   2
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   87
            Top             =   8340
            Width           =   4455
         End
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
            Index           =   1
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   2220
            Width           =   4455
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
            Index           =   1
            Left            =   1860
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1620
            Width           =   4455
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de DESCENTE du NIVEAU INTERMEDIAIRE au NIVEAU BAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   54
            Left            =   6960
            TabIndex        =   108
            Top             =   8640
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "DIVERS"
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
            Height          =   255
            Index           =   55
            Left            =   660
            TabIndex        =   107
            Top             =   9900
            Width           =   5775
         End
         Begin VB.Label LTempsMonteeBasVersHaut 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   7020
            TabIndex        =   106
            Top             =   11460
            Width           =   6135
         End
         Begin VB.Label LTempsMonteeBasVersIntermediaire 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   7020
            TabIndex        =   105
            Top             =   10620
            Width           =   6135
         End
         Begin VB.Label LTempsDescenteIntermediaireVersBas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   7020
            TabIndex        =   104
            Top             =   8940
            Width           =   6135
         End
         Begin VB.Label LTempsDescenteHautVersBas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   7020
            TabIndex        =   103
            Top             =   8100
            Width           =   6135
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MONTEE du NIVEAU BAS au NIVEAU HAUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   53
            Left            =   6960
            TabIndex        =   102
            Top             =   11160
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MONTEE du NIVEAU BAS au NIVEAU INTERMEDIAIRE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   51
            Left            =   6960
            TabIndex        =   101
            Top             =   10320
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de DESCENTE du NIVEAU HAUT au NIVEAU BAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   50
            Left            =   6960
            TabIndex        =   100
            Top             =   7800
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de DESCENTE des ACCROCHES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   49
            Left            =   780
            TabIndex        =   99
            Top             =   11160
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MONTEE des ACCROCHES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   48
            Left            =   780
            TabIndex        =   98
            Top             =   10380
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS du DEPLACEMENT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   47
            Left            =   780
            TabIndex        =   97
            Top             =   8820
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Poste d'ARRIVEE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   46
            Left            =   780
            TabIndex        =   96
            Top             =   8280
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Poste de DEPART"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   45
            Left            =   780
            TabIndex        =   95
            Top             =   7740
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "LEVAGE"
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
            Height          =   255
            Index           =   44
            Left            =   6840
            TabIndex        =   94
            Top             =   7320
            Width           =   6495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSLATION"
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
            Height          =   255
            Index           =   43
            Left            =   660
            TabIndex        =   93
            Top             =   7320
            Width           =   5775
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MOUVEMENTS sur le PONT 2"
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
            Height          =   255
            Index           =   29
            Left            =   600
            TabIndex        =   92
            Top             =   6720
            Width           =   12795
         End
         Begin VB.Label LTempsMonteeAccroches 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   91
            Top             =   10680
            Width           =   5415
         End
         Begin VB.Label LTempsDescenteAccroches 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   90
            Top             =   11460
            Width           =   5415
         End
         Begin VB.Label LTempsTranslation 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   89
            Top             =   9120
            Width           =   5415
         End
         Begin VB.Label LTempsTranslation 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   8
            Top             =   3000
            Width           =   5415
         End
         Begin VB.Label LTempsDescenteAccroches 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   18
            Top             =   5340
            Width           =   5415
         End
         Begin VB.Label LTempsMonteeAccroches 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   19
            Top             =   4560
            Width           =   5415
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   13
            Left            =   17820
            TabIndex        =   73
            Top             =   3060
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   12
            Left            =   17820
            TabIndex        =   72
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   11
            Left            =   17820
            TabIndex        =   71
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   10
            Left            =   17820
            TabIndex        =   70
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   17220
            TabIndex        =   69
            Top             =   3060
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   17220
            TabIndex        =   68
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   17220
            TabIndex        =   67
            Top             =   2220
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
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
            Left            =   17220
            TabIndex        =   66
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   9
            Left            =   15480
            TabIndex        =   65
            Top             =   5160
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   8
            Left            =   15480
            TabIndex        =   64
            Top             =   4740
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   7
            Left            =   15480
            TabIndex        =   63
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   6
            Left            =   15480
            TabIndex        =   62
            Top             =   3900
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   5
            Left            =   15480
            TabIndex        =   61
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   4
            Left            =   15480
            TabIndex        =   60
            Top             =   3060
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   3
            Left            =   15480
            TabIndex        =   59
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   15480
            TabIndex        =   58
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   9
            Left            =   14880
            TabIndex        =   57
            Top             =   5160
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   8
            Left            =   14880
            TabIndex        =   56
            Top             =   4740
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   7
            Left            =   14880
            TabIndex        =   55
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   6
            Left            =   14880
            TabIndex        =   54
            Top             =   3900
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   5
            Left            =   14880
            TabIndex        =   53
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   4
            Left            =   14880
            TabIndex        =   52
            Top             =   3060
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   3
            Left            =   14880
            TabIndex        =   51
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   2
            Left            =   14880
            TabIndex        =   50
            Top             =   2220
            Width           =   495
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   15480
            TabIndex        =   49
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de FERMETURE des COUVERCLES"
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
            Height          =   255
            Index           =   0
            Left            =   20220
            TabIndex        =   48
            Top             =   1260
            Width           =   4695
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS d'OUVERTURE des COUVERCLES"
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
            Height          =   255
            Index           =   100
            Left            =   14640
            TabIndex        =   47
            Top             =   1260
            Width           =   4695
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   1
            Left            =   14880
            TabIndex        =   46
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MOUVEMENTS sur les CUVES"
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
            Height          =   255
            Index           =   27
            Left            =   14280
            TabIndex        =   33
            Top             =   600
            Width           =   10995
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   56
            Left            =   17220
            TabIndex        =   31
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label LTempsOuvertureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   14
            Left            =   17820
            TabIndex        =   30
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MOUVEMENTS sur le PONT 1"
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
            Height          =   255
            Index           =   28
            Left            =   600
            TabIndex        =   27
            Top             =   600
            Width           =   12795
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "TRANSLATION"
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
            Height          =   255
            Index           =   30
            Left            =   660
            TabIndex        =   26
            Top             =   1200
            Width           =   5775
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "LEVAGE"
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
            Height          =   255
            Index           =   32
            Left            =   6840
            TabIndex        =   25
            Top             =   1200
            Width           =   6495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Poste de DEPART"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   33
            Left            =   780
            TabIndex        =   24
            Top             =   1620
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Poste d'ARRIVEE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   34
            Left            =   780
            TabIndex        =   23
            Top             =   2160
            Width           =   975
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS du DEPLACEMENT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   780
            TabIndex        =   22
            Top             =   2700
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MONTEE des ACCROCHES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   780
            TabIndex        =   21
            Top             =   4260
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de DESCENTE des ACCROCHES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   780
            TabIndex        =   20
            Top             =   5040
            Width           =   5535
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de DESCENTE du NIVEAU HAUT au NIVEAU BAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   6960
            TabIndex        =   17
            Top             =   1680
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MONTEE du NIVEAU BAS au NIVEAU INTERMEDIAIRE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   6960
            TabIndex        =   16
            Top             =   4200
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de MONTEE du NIVEAU BAS au NIVEAU HAUT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   6960
            TabIndex        =   15
            Top             =   5040
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TEMPS de DESCENTE du NIVEAU INTERMEDIAIRE au NIVEAU BAS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   6960
            TabIndex        =   14
            Top             =   2520
            Width           =   6255
            WordWrap        =   -1  'True
         End
         Begin VB.Label LTempsDescenteHautVersBas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   7020
            TabIndex        =   13
            Top             =   1980
            Width           =   6135
         End
         Begin VB.Label LTempsDescenteIntermediaireVersBas 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   7020
            TabIndex        =   12
            Top             =   2820
            Width           =   6135
         End
         Begin VB.Label LTempsMonteeBasVersIntermediaire 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   7020
            TabIndex        =   11
            Top             =   4500
            Width           =   6135
         End
         Begin VB.Label LTempsMonteeBasVersHaut 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   7020
            TabIndex        =   10
            Top             =   5340
            Width           =   6135
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            BackStyle       =   0  'Transparent
            Caption         =   "DIVERS"
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
            Height          =   255
            Index           =   31
            Left            =   660
            TabIndex        =   9
            Top             =   3780
            Width           =   5775
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   32
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   540
            Width           =   13035
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   36
            Left            =   600
            Top             =   1140
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   37
            Left            =   6780
            Top             =   1140
            Width           =   6615
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   38
            Left            =   600
            Top             =   3720
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   42
            Left            =   780
            Shape           =   4  'Rounded Rectangle
            Top             =   4500
            Width           =   5535
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   43
            Left            =   780
            Shape           =   4  'Rounded Rectangle
            Top             =   5280
            Width           =   5535
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   1755
            Index           =   41
            Left            =   600
            Top             =   4080
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   45
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   1920
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   46
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   2760
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   47
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   4440
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   49
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   5280
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   4335
            Index           =   44
            Left            =   6780
            Top             =   1500
            Width           =   6615
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   34
            Left            =   480
            Shape           =   4  'Rounded Rectangle
            Top             =   6660
            Width           =   13035
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   35
            Left            =   600
            Top             =   7260
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   50
            Left            =   6780
            Top             =   7260
            Width           =   6615
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   51
            Left            =   600
            Top             =   9840
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   61
            Left            =   780
            Shape           =   4  'Rounded Rectangle
            Top             =   9060
            Width           =   5535
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   1995
            Index           =   62
            Left            =   600
            Top             =   7620
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   52
            Left            =   780
            Shape           =   4  'Rounded Rectangle
            Top             =   10620
            Width           =   5535
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   53
            Left            =   780
            Shape           =   4  'Rounded Rectangle
            Top             =   11400
            Width           =   5535
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   1755
            Index           =   54
            Left            =   600
            Top             =   10200
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   55
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   8040
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   56
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   8880
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   57
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   10560
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   59
            Left            =   6960
            Shape           =   4  'Rounded Rectangle
            Top             =   11400
            Width           =   6255
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   4335
            Index           =   60
            Left            =   6780
            Top             =   7620
            Width           =   6615
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            Height          =   5835
            Index           =   63
            Left            =   300
            Top             =   6420
            Width           =   13395
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   31
            Left            =   14160
            Shape           =   4  'Rounded Rectangle
            Top             =   540
            Width           =   11235
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   29
            Left            =   14520
            Top             =   1200
            Width           =   4935
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   0
            Left            =   20100
            Top             =   1200
            Width           =   4935
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   26
            Left            =   20460
            TabIndex        =   86
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   25
            Left            =   20460
            TabIndex        =   85
            Top             =   2220
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   24
            Left            =   20460
            TabIndex        =   84
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   23
            Left            =   20460
            TabIndex        =   83
            Top             =   3060
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   22
            Left            =   20460
            TabIndex        =   82
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   21
            Left            =   20460
            TabIndex        =   81
            Top             =   3900
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   20
            Left            =   20460
            TabIndex        =   80
            Top             =   4320
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   19
            Left            =   20460
            TabIndex        =   79
            Top             =   4740
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   18
            Left            =   20460
            TabIndex        =   78
            Top             =   5160
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   17
            Left            =   22740
            TabIndex        =   77
            Top             =   1800
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   16
            Left            =   22740
            TabIndex        =   76
            Top             =   2220
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   15
            Left            =   22740
            TabIndex        =   75
            Top             =   2640
            Width           =   495
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   240
            Index           =   14
            Left            =   22740
            TabIndex        =   74
            Top             =   3060
            Width           =   495
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   1
            Left            =   21060
            TabIndex        =   45
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   2
            Left            =   21060
            TabIndex        =   44
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   3
            Left            =   21060
            TabIndex        =   43
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   4
            Left            =   21060
            TabIndex        =   42
            Top             =   3060
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   5
            Left            =   21060
            TabIndex        =   41
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   7
            Left            =   21060
            TabIndex        =   40
            Top             =   4320
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   8
            Left            =   21060
            TabIndex        =   39
            Top             =   4740
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   9
            Left            =   21060
            TabIndex        =   38
            Top             =   5160
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   10
            Left            =   23340
            TabIndex        =   37
            Top             =   1800
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   11
            Left            =   23340
            TabIndex        =   36
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   12
            Left            =   23340
            TabIndex        =   35
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   13
            Left            =   23340
            TabIndex        =   34
            Top             =   3060
            Width           =   1335
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   6
            Left            =   21060
            TabIndex        =   32
            Top             =   3900
            Width           =   1335
         End
         Begin VB.Label LLibelles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BackStyle       =   0  'Transparent
            Caption         =   "-"
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
            Height          =   240
            Index           =   57
            Left            =   22740
            TabIndex        =   29
            Top             =   3480
            Width           =   495
         End
         Begin VB.Label LTempsFermetureCouvercles 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
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
            Height          =   255
            Index           =   14
            Left            =   23340
            TabIndex        =   28
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   9
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   5100
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   8
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   7
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   4260
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   6
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   5
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   4
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   3
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   2
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   1
            Left            =   21000
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   10
            Left            =   23280
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   11
            Left            =   23280
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   12
            Left            =   23280
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   13
            Left            =   23280
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Shape SDecorationFermetureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   14
            Left            =   23280
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   1455
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   26
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   25
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   24
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   23
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   22
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   21
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   3840
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   20
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   4260
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   19
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   4680
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   18
            Left            =   20340
            Shape           =   4  'Rounded Rectangle
            Top             =   5100
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   17
            Left            =   22620
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   16
            Left            =   22620
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   15
            Left            =   22620
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00808000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   14
            Left            =   22620
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   65
            Left            =   22620
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   675
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   14
            Left            =   17760
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   13
            Left            =   17760
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   12
            Left            =   17760
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   11
            Left            =   17760
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   10
            Left            =   17760
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   64
            Left            =   17100
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00808000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   13
            Left            =   17100
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   12
            Left            =   17100
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   11
            Left            =   17100
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   10
            Left            =   17100
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   675
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   1
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   2
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   3
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   4
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   5
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   6
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   3840
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   7
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   4260
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   8
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   4680
            Width           =   1455
         End
         Begin VB.Shape SDecorationOuvertureCouvercles 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   9
            Left            =   15420
            Shape           =   4  'Rounded Rectangle
            Top             =   5100
            Width           =   1455
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   1
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   1740
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   2
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   3
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   2580
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   4
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   3000
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FF00FF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   5
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   3420
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   6
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   3840
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H0000FF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   7
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   4260
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFF00&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   8
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   4680
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   9
            Left            =   14760
            Shape           =   4  'Rounded Rectangle
            Top             =   5100
            Width           =   675
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   4095
            Index           =   27
            Left            =   14520
            Top             =   1560
            Width           =   4935
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   4095
            Index           =   28
            Left            =   20100
            Top             =   1560
            Width           =   4935
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            Height          =   5835
            Index           =   30
            Left            =   13980
            Top             =   300
            Width           =   11595
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   39
            Left            =   780
            Shape           =   4  'Rounded Rectangle
            Top             =   2940
            Width           =   5535
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   1  'Opaque
            Height          =   1995
            Index           =   40
            Left            =   600
            Top             =   1500
            Width           =   5895
         End
         Begin VB.Shape SDecoration 
            BackColor       =   &H00C0C000&
            BackStyle       =   1  'Opaque
            Height          =   5835
            Index           =   33
            Left            =   300
            Top             =   300
            Width           =   13395
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
      ScaleWidth      =   16020
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   13995
      Width           =   16080
      Begin VB.PictureBox PBOutilsDeplacementFenetre 
         BackColor       =   &H00E0E0E0&
         Height          =   1035
         Left            =   0
         ScaleHeight     =   975
         ScaleWidth      =   1155
         TabIndex        =   110
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
         Begin VB.HScrollBar HSDeplacementFenetre 
            Height          =   255
            LargeChange     =   300
            Left            =   0
            SmallChange     =   100
            TabIndex        =   113
            Top             =   720
            Width           =   915
         End
         Begin VB.VScrollBar VSDeplacementFenetre 
            Height          =   975
            LargeChange     =   300
            Left            =   900
            SmallChange     =   100
            TabIndex        =   112
            Top             =   0
            Width           =   255
         End
         Begin VB.CommandButton CBAgrandirFenetre 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Agrandir"
            DownPicture     =   "FTempsMouvements.frx":0442
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
            Picture         =   "FTempsMouvements.frx":05EC
            Style           =   1  'Graphical
            TabIndex        =   111
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
         DownPicture     =   "FTempsMouvements.frx":0796
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
         Left            =   14580
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTempsMouvements.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBEnregistrement 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Enregistrement des valeurs"
         DownPicture     =   "FTempsMouvements.frx":159A
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
         Left            =   11280
         MaskColor       =   &H00FF00FF&
         Picture         =   "FTempsMouvements.frx":1C9C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   " Enregistrement des valeurs "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   3195
      End
      Begin VB.Timer TimerEtatsTempsMouvements 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1500
         Top             =   180
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   2340
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      Picture         =   "FTempsMouvements.frx":239E
      ScaleHeight     =   315
      ScaleWidth      =   16020
      TabIndex        =   0
      Top             =   0
      Width           =   16080
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPS DE MOUVEMENTS"
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
         TabIndex        =   1
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "FTempsMouvements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre affichant les temps de mouvements
' Nom                    : FTempsMouvements.frm
' Date de création : 15/12/2010
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "TEMPS DE MOUVEMENTS"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum ONGLETS
    O_TEMPS_MOUVEMENTS_CUVES = 0
    O_TEMPS_MOUVEMENTS_PONT_1 = 1
    O_TEMPS_MOUVEMENTS_PONT_2 = 2
End Enum

'--- variables privées ---
Private PremiereActivation As Boolean

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                             'numéro de la fenêtre lorsqu'elle devient active
    
Private Sub CBAgrandirFENETRE_Click()
    On Error Resume Next
    Me.WindowState = vbMaximized
End Sub

Private Sub CBEnregistrement_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- gestion de la souris ---
    SourisEnAttente True
    
    '--- lancement de l'enregistrement des valeurs ---
    Bidon = EnregistrementTempsMouvements
    
    '--- rafraichir l'intégralité des valeurs ---
    Bidon = ChargeTempsMouvements

    '--- gestion de la souris ---
    SourisEnAttente False

End Sub

Private Sub CBEnregistrement_GotFocus()
    
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

Private Sub CBQuitter_Click()
    On Error Resume Next
    
    
    'insertionClipperPointage 11
    
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

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        PremiereActivation = True
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

Private Sub LTempsDescenteAccroches_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsDescenteAccroches(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de DESCENTE des ACCROCHES pour ce PONT")
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TempsAccrochesChargeVersBas = CSng(Reponse)
            EnregistrementTempsMouvements
        End If
    End If

End Sub

Private Sub LTempsDescenteHautVersBas_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsDescenteHautVersBas(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de DESCENTE du NIVEAU HAUT vers le NIVEAU BAS pour ce PONT")
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TempsDescenteHautVersBas = CSng(Reponse)
            EnregistrementTempsMouvements
        End If
    End If

End Sub

Private Sub LTempsDescenteIntermediaireVersBas_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsDescenteIntermediaireVersBas(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de DESCENTE du NIVEAU INTERMEDIAIRE vers le NIVEAU BAS pour ce PONT")
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TempsDescenteIntermediaireVersBas = CSng(Reponse)
            EnregistrementTempsMouvements
        End If
    End If

End Sub

Private Sub LTempsFermetureCouvercles_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsFermetureCouvercles(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de FERMETURE de ces COUVERCLES")
        If IsNumeric(Reponse) = True Then
            TEtatsCuves(Index).TTempsMouvements.TempsFermetureCouvercles = CSng(Reponse)
        End If
    End If

End Sub

Private Sub LTempsMonteeAccroches_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsMonteeAccroches(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de MONTEE des ACCROCHES pour ce PONT")
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TempsAccrochesChargeVersHaut = CSng(Reponse)
            EnregistrementTempsMouvements
        End If
    End If

End Sub

Private Sub LTempsMonteeBasVersHaut_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsMonteeBasVersHaut(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de MONTEE du NIVEAU BAS vers le NIVEAU HAUT pour ce PONT")
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TempsMonteeBasVersHaut = CSng(Reponse)
            EnregistrementTempsMouvements
        End If
    End If

End Sub

Private Sub LTempsMonteeBasVersIntermediaire_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsMonteeBasVersIntermediaire(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps de MONTEE du NIVEAU BAS vers le NIVEAU INTERMEDIAIRE pour ce PONT")
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TempsMonteeBasVersIntermediaire = CSng(Reponse)
            EnregistrementTempsMouvements
        End If
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

Private Sub LTempsOuvertureCouvercles_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim Reponse As String
        
    '--- analyse de la réponse ---
    If LTempsOuvertureCouvercles(Index).BackColor = COULEURS.BLANC Then
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & "au temps d'OUVERTURE de ces COUVERCLES")
        If IsNumeric(Reponse) = True Then
            TEtatsCuves(Index).TTempsMouvements.TempsOuvertureCouvercles = CSng(Reponse)
        End If
    End If

End Sub

Private Sub LTempsTranslation_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    Dim NumPosteDepart As Integer, _
            NumPosteArrivee As Integer
    Dim Reponse As String

    '--- affectation ---
    NumPosteDepart = CBNumPosteDepart(Index).ListIndex
    NumPosteArrivee = CBNumPosteArrivee(Index).ListIndex

    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE Then

        '--- appel de la boite de dialogues ---
        Reponse = InputBox("Entrez la valeur numérique sans unité correspondant" & vbCrLf & _
                                          "au temps de DEPLACEMENT en TRANSLATION" & vbCrLf & _
                                          "du POSTE " & TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                          " au POSTE " & TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                          " pour ce PONT")
        
        '--- affectation de la valeur ---
        If IsNumeric(Reponse) = True Then
            TEtatsPonts(Index).TTempsMouvements.TTempsTranslation(NumPosteDepart, NumPosteArrivee) = CSng(Reponse)
            EnregistrementTempsMouvements
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

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBEnregistrement.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBEnregistrement.Width
    
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
' Rôle      : Analyse les changements d'états
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AnalyseChangementsEtats()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
            
    '--- constantes privées ---
    Const TEXTE_SECONDES As String = " secondes"
    Const TEXTE_EXTRAPOLATION As String = " (par extrapolation)"
    
    '--- déclaration ---
    Dim a As Integer, _
            NumPoste As Integer, _
            NumPosteDepart As Integer, _
            NumPosteArrivee As Integer, _
            IdxPontPourExtrapolation As Integer
    Dim Texte As String
    
    '********************************************************************************************************************
    '                                                                               CUVES
    '********************************************************************************************************************
   
    '--- pour les cuves ---
    For a = LBound(TEtatsCuves()) To UBound(TEtatsCuves())

        With TEtatsCuves(a)
            
            '--- recherche du poste pour l'agitation de la charge et les couvercles ---
            NumPoste = CorrespondanceCuvesAPIPostes(a)

            '--- couvercles ---
            If NumPoste > 0 Then

                If TEtatsPostes(NumPoste).DefinitionPoste.PresenceCouvercles = True Then

                    '--- ouverture des couvercles ---
                    If .TTempsMouvements.TempsOuvertureCouvercles = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsOuvertureCouvercles & TEXTE_SECONDES
                    End If
                    
                    '---affichage du texte ---
                    AffichageTexte LTempsOuvertureCouvercles(a), Texte
                    
                    '--- fermeture des couvercles ---
                    If .TTempsMouvements.TempsFermetureCouvercles = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsFermetureCouvercles & TEXTE_SECONDES
                    End If
                    
                    '---affichage du texte ---
                    AffichageTexte LTempsFermetureCouvercles(a), Texte
                
                End If
            
            End If
        
        End With

    Next a

    '********************************************************************************************************************
    '                                                                               PONTS
    '********************************************************************************************************************
    
    '--- pour les ponts ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
            
        With TEtatsPonts(a)
            
            '--- affectation ---
            NumPosteDepart = CBNumPosteDepart(a).ListIndex
            NumPosteArrivee = CBNumPosteArrivee(a).ListIndex
            
            '--- générer l'index du pont pour l'extrapolation ---
            'l'extrapolation permet de récupérer la mesure sur le pont 2 si la mesure d'un temps est égal à 0
            'sur le pont 1 et inversement
            'si le temps est égal à 0 sur les 2 ponts alors pas d'affichage (le mouvement n'a jamais été fait ou n'est
            'pas nécessaire au fonctionnement de l'installation)
            Select Case a
                Case PONTS.P_1: IdxPontPourExtrapolation = PONTS.P_2
                Case PONTS.P_2: IdxPontPourExtrapolation = PONTS.P_1
                Case Else
            End Select
            
            '**********************************************************************************************************
            '*                                                                translation
            '**********************************************************************************************************
            If NumPosteDepart > 0 And NumPosteArrivee > 0 Then
                    
                If .TTempsMouvements.TTempsTranslation(NumPosteDepart, NumPosteArrivee) = 0 Then
                
                    '--- vérifier si il existe un temps sur l'autre pont ---
                    With TEtatsPonts(IdxPontPourExtrapolation)
                        If .TTempsMouvements.TTempsTranslation(NumPosteDepart, NumPosteArrivee) = 0 Then
                            Texte = ""
                        Else
                            Texte = .TTempsMouvements.TTempsTranslation(NumPosteDepart, NumPosteArrivee) & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                        End If
                    End With
                
                Else
                
                    '--- affectation du temps de translation ---
                    Texte = .TTempsMouvements.TTempsTranslation(NumPosteDepart, NumPosteArrivee) & TEXTE_SECONDES
                
                End If
                    
            Else
                
                '--- affectation ---
                Texte = ""
            
            End If
             
            '---affichage du texte ---
            AffichageTexte LTempsTranslation(a), Texte
            
            '**********************************************************************************************************
            '*                                                              Montée des accroches
            '**********************************************************************************************************
            If .TTempsMouvements.TempsAccrochesChargeVersHaut = 0 Then
                
                '--- vérifier si il existe un temps sur l'autre pont ---
                With TEtatsPonts(IdxPontPourExtrapolation)
                    If .TTempsMouvements.TempsAccrochesChargeVersHaut = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsAccrochesChargeVersHaut & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                    End If
                End With
                
            Else
                Texte = .TTempsMouvements.TempsAccrochesChargeVersHaut & TEXTE_SECONDES
            End If
            
            '---affichage du texte ---
            AffichageTexte LTempsMonteeAccroches(a), Texte
            
            '**********************************************************************************************************
            '*                                                            Descente des accroches
            '**********************************************************************************************************
            If .TTempsMouvements.TempsAccrochesChargeVersBas = 0 Then
                
                '--- vérifier si il existe un temps sur l'autre pont ---
                With TEtatsPonts(IdxPontPourExtrapolation)
                    If .TTempsMouvements.TempsAccrochesChargeVersBas = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsAccrochesChargeVersBas & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                    End If
                End With
            
            Else
                Texte = .TTempsMouvements.TempsAccrochesChargeVersBas & TEXTE_SECONDES
            End If
            
            '---affichage du texte ---
            AffichageTexte LTempsDescenteAccroches(a), Texte
    
            '**********************************************************************************************************
            '*                                   temps de descente du niveau haut au niveau bas
            '**********************************************************************************************************
            If .TTempsMouvements.TempsDescenteHautVersBas = 0 Then
                
                '--- vérifier si il existe un temps sur l'autre pont ---
                With TEtatsPonts(IdxPontPourExtrapolation)
                    If .TTempsMouvements.TempsDescenteHautVersBas = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsDescenteHautVersBas & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                    End If
                End With
            
            Else
                Texte = .TTempsMouvements.TempsDescenteHautVersBas & TEXTE_SECONDES
            End If
            
            '---affichage du texte ---
            AffichageTexte LTempsDescenteHautVersBas(a), Texte
            
            '**********************************************************************************************************
            '*                                temps de descente du niveau intermédiaire au niveau bas
            '**********************************************************************************************************
            If .TTempsMouvements.TempsDescenteIntermediaireVersBas = 0 Then
                
                '--- vérifier si il existe un temps sur l'autre pont ---
                With TEtatsPonts(IdxPontPourExtrapolation)
                    If .TTempsMouvements.TempsDescenteIntermediaireVersBas = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsDescenteIntermediaireVersBas & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                    End If
                End With
            
            Else
                Texte = .TTempsMouvements.TempsDescenteIntermediaireVersBas & TEXTE_SECONDES
            End If
            
            '---affichage du texte ---
            AffichageTexte LTempsDescenteIntermediaireVersBas(a), Texte
            
            '**********************************************************************************************************
            '*                                  temps de montée du niveau bas au niveau intermédiaire
            '**********************************************************************************************************
            If .TTempsMouvements.TempsMonteeBasVersIntermediaire = 0 Then
                
                '--- vérifier si il existe un temps sur l'autre pont ---
                With TEtatsPonts(IdxPontPourExtrapolation)
                    If .TTempsMouvements.TempsMonteeBasVersIntermediaire = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsMonteeBasVersIntermediaire & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                    End If
                End With
            
            Else
                Texte = .TTempsMouvements.TempsMonteeBasVersIntermediaire & TEXTE_SECONDES
            End If
            
            '---affichage du texte ---
            AffichageTexte LTempsMonteeBasVersIntermediaire(a), Texte
            
            '**********************************************************************************************************
            '*                                       temps de montée du niveau bas au niveau haut
            '**********************************************************************************************************
            If .TTempsMouvements.TempsMonteeBasVersHaut = 0 Then
                
                '--- vérifier si il existe un temps sur l'autre pont ---
                With TEtatsPonts(IdxPontPourExtrapolation)
                    If .TTempsMouvements.TempsMonteeBasVersHaut = 0 Then
                        Texte = ""
                    Else
                        Texte = .TTempsMouvements.TempsMonteeBasVersHaut & TEXTE_SECONDES & TEXTE_EXTRAPOLATION
                    End If
                End With
            
            Else
                Texte = .TTempsMouvements.TempsMonteeBasVersHaut & TEXTE_SECONDES
            End If
            
            '---affichage du texte ---
            AffichageTexte LTempsMonteeBasVersHaut(a), Texte
    
        End With
    
    Next a

End Sub

Private Sub TimerEtatsTempsMouvements_Timer()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- appel de la routine ---
    TimerEtatsTempsMouvements.Enabled = False
    AnalyseChangementsEtats
    TimerEtatsTempsMouvements.Enabled = True

    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

Private Sub VSDeplacementFENETRE_Change()
    On Error Resume Next
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Top = -VSDeplacementFenetre.value
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
    With TimerEtatsTempsMouvements
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFTempsMouvements = Nothing

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
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
           b As Integer

    '--- affectation ---
  
    '--- divers sur la fenêtre ---
    With Me
        .Caption = TITRE_FENETRE
        .WindowState = vbMaximized
    End With
    PBDeplacementFenetre(ZONES_DEPLACEMENT_FENETRE.Z_FILLE).Picture = ImgFondEspace
    PBBoutons.Picture = ImgFondDesBoutons
        
    '--- gestion des détails ---

    '--- transfert des postes de départ et arrivée ---
    For a = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
        With CBNumPosteDepart(a)
            .Clear
            .AddItem ("")
            .ItemData(.NewIndex) = 0
        End With
        With CBNumPosteArrivee(a)
            .Clear
            .AddItem ("")
            .ItemData(.NewIndex) = 0
        End With
    Next a
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        For b = LBound(TEtatsPonts()) To UBound(TEtatsPonts())
            With TEtatsPostes(a).DefinitionPoste
                CBNumPosteDepart(b).AddItem (.NomPoste & " - " & .LibellePoste)
                CBNumPosteDepart(b).ItemData(CBNumPosteDepart(b).NewIndex) = .NumPoste
                CBNumPosteArrivee(b).AddItem (.NomPoste & " - " & .LibellePoste)
                CBNumPosteArrivee(b).ItemData(CBNumPosteArrivee(b).NewIndex) = .NumPoste
            End With
        Next b
    Next a
    
    '--- affectation ---

    '--- affiche les élements ayant des temps de mouvements ---
    AffichageElementsAyantTempsMouvements
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche les élements ayant des temps de mouvements
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AffichageElementsAyantTempsMouvements()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim a As Integer, _
            NumPoste As Integer
    
    '--- affichage pour les cuves ---
    For a = LBound(TEtatsCuves()) To UBound(TEtatsCuves())

        With TEtatsCuves(a)

            '--- recherche du poste pour les couvercles ---
            NumPoste = CorrespondanceCuvesAPIPostes(a)

            '--- couvercles ---
            If NumPoste > 0 Then
                With TEtatsPostes(NumPoste)
                    If .DefinitionPoste.PresenceCouvercles = True Then
                        'LTempsOuvertureCouvercles(a).Visible = True
                        'SDecorationOuvertureCouvercles(a).BackColor = COULEURS.BLANC
                        'LTempsFermetureCouvercles(a).Visible = True
                        'SDecorationFermetureCouvercles(a).BackColor = COULEURS.BLANC
                    Else
                        'LTempsOuvertureCouvercles(a).Visible = False
                        'SDecorationOuvertureCouvercles(a).BackColor = COULEURS.GRIS_1
                        'LTempsFermetureCouvercles(a).Visible = False
                        'SDecorationFermetureCouvercles(a).BackColor = COULEURS.GRIS_1
                    End If
                End With
            End If

        End With

    Next a

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la fenêtre
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- analyse des changements d'états ---
    AnalyseChangementsEtats

    '--- lancement du timer ---
    TimerEtatsTempsMouvements.Enabled = True
                
End Sub




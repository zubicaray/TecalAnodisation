VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FChoixImpression 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choix de l'impression"
   ClientHeight    =   6660
   ClientLeft      =   5265
   ClientTop       =   3990
   ClientWidth     =   12810
   Icon            =   "FChoixImpression.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FPersonneEmettrice 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Personne émettrice"
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
      TabIndex        =   49
      Top             =   120
      Width           =   12435
      Begin MSDataListLib.DataCombo DCPersonneEmettrice 
         Bindings        =   "FChoixImpression.frx":014A
         Height          =   360
         Left            =   180
         TabIndex        =   50
         Top             =   300
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "NomComplet"
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   "PersonnesEmettrices"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame FMarges 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Marges en millimètres"
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
      Left            =   8340
      TabIndex        =   42
      Top             =   1140
      Width           =   4275
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   8
         Left            =   3240
         TabIndex        =   41
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   7
         Left            =   3240
         TabIndex        =   37
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   6
         Left            =   3240
         TabIndex        =   33
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   5
         Left            =   3240
         TabIndex        =   29
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   3240
         TabIndex        =   25
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   3240
         TabIndex        =   21
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   3240
         TabIndex        =   17
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TBMargesBasse 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   3240
         TabIndex        =   13
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   8
         Left            =   2220
         TabIndex        =   40
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   7
         Left            =   2220
         TabIndex        =   36
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   6
         Left            =   2220
         TabIndex        =   32
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Index           =   5
         Left            =   2220
         TabIndex        =   28
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   2220
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   2220
         TabIndex        =   20
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   2220
         TabIndex        =   16
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TBMargesDroite 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   2220
         TabIndex        =   12
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   1200
         TabIndex        =   39
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   1200
         TabIndex        =   35
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   1200
         TabIndex        =   31
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   1200
         TabIndex        =   27
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   23
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   19
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
         Alignment       =   1  'Right Justify
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
         Left            =   1200
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TBMargesHaute 
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
         Left            =   1200
         TabIndex        =   11
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Index           =   8
         Left            =   180
         TabIndex        =   38
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Index           =   7
         Left            =   180
         TabIndex        =   34
         Top             =   3300
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Index           =   6
         Left            =   180
         TabIndex        =   30
         Top             =   2880
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Index           =   5
         Left            =   180
         TabIndex        =   26
         Top             =   2460
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "###0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
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
         Left            =   180
         TabIndex        =   10
         Top             =   780
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         TabIndex        =   14
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         TabIndex        =   18
         Top             =   1620
         Width           =   855
      End
      Begin VB.TextBox TBMargesGauche 
         Alignment       =   1  'Right Justify
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
         Left            =   180
         TabIndex        =   22
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Basse"
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
         Left            =   3180
         TabIndex        =   48
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Droite"
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
         Left            =   2160
         TabIndex        =   47
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gauche"
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
         Index           =   0
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Haute"
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
         Left            =   1140
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   12750
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   5565
      Width           =   12810
      Begin VB.CommandButton CBAnnuler 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Annuler"
         DownPicture     =   "FChoixImpression.frx":015B
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
         Picture         =   "FChoixImpression.frx":085D
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   " Annuler les dernières modifications "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBValider 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Valider"
         DownPicture     =   "FChoixImpression.frx":0F5F
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
         Left            =   5700
         MaskColor       =   &H00FF00FF&
         Picture         =   "FChoixImpression.frx":1661
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   " Valider l'enregistrement "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Timer TimerImprimanteParDefaut 
         Interval        =   1000
         Left            =   1860
         Top             =   180
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   345
         Left            =   600
         Top             =   150
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.Frame FOptionsImpression 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Options d'impression"
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
      Left            =   180
      TabIndex        =   0
      Top             =   1140
      Width           =   7935
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Index           =   8
         Left            =   180
         TabIndex        =   9
         Top             =   3720
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Index           =   7
         Left            =   180
         TabIndex        =   8
         Top             =   3300
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Index           =   6
         Left            =   180
         TabIndex        =   7
         Top             =   2880
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   2460
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Left            =   180
         TabIndex        =   5
         Top             =   2040
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Left            =   180
         TabIndex        =   4
         Top             =   1620
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Left            =   180
         TabIndex        =   3
         Top             =   1200
         Width           =   7575
      End
      Begin VB.OptionButton OBOptionsImpression 
         BackColor       =   &H00C0E0FF&
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
         Left            =   180
         TabIndex        =   2
         Top             =   780
         Width           =   7575
      End
      Begin VB.Label LImprimanteParDefaut 
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   1500
         TabIndex        =   46
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label LLibelles 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Imprimante"
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
         Index           =   2
         Left            =   180
         TabIndex        =   45
         Top             =   360
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FChoixImpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre affichant le choix de l'impression
' Nom                    : FChoixImpression.frm
' Date de création : 15/11/1999
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const TITRE_FENETRE As String = "Choix de l'impression"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- variables privées ---
Private PremiereActivation As Boolean
Private ChangementMarges As Boolean         'indique un changement des marges

'--- variables publiques ---
Public OptionSelectionnee As Integer              'option sélectionnée
Public NumFenetre As Long                              'numéro de la fenêtre lorsqu'elle devient active
Public NumfenetreAppel As Long                     'numéro de fenetre ayant lancé l'appel

Private Sub CBAnnuler_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- affectation ---
    OptionSelectionnee = 0
    
    '--- déchargement de la fenêtre ---
    DechargeFenetre

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

Private Sub CBValider_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- sauvegarde des valeurs ---
    EnregistrementMarges
    CalculMargesImpressionTwips
    EnregistrementPersonneEmettrice
    
    '--- déchargement de la fenêtre ---
    DechargeFenetre
    
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

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- centrage de la fenetre ---
    Centrefenetre Me
    
    '--- renseigne la fenêtre principale ---
    NumFenetre = FENETRES.F_CHOIX_IMPRESSION
    RenseigneFPrincipale
    
    If PremiereActivation = False Then
    
        '--- divers chargements ---
        ChargementMarges
        ChargementPersonneEmettrice
    
        '--- recherche de l'imprimante par défaut ---
        TimerImprimanteParDefaut_Timer
    
        '--- affectation ---
        ChangementMarges = False
        PremiereActivation = True
        
        '--- placement du focus ---
        If OBOptionsImpression(OptionSelectionnee).Visible = True Then OBOptionsImpression(OptionSelectionnee).SetFocus
    
    End If

End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- images des fonds ---
    'Me.Picture = ImgFondOrange2
    PBBoutons.Picture = ImgFondDesBoutons

End Sub

Private Sub OBOptionsImpression_Click(Index As Integer)
    On Error Resume Next
    OptionSelectionnee = Index
End Sub

Private Sub OBOptionsImpression_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then CBValider_Click
End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBAnnuler.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBAnnuler.Width
    CBValider.Left = CBAnnuler.Left - MARGES.M_ENTRE_BOUTONS - CBAnnuler.Width
    
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

Private Sub TBMargesBasse_Change(Index As Integer)
    On Error Resume Next
    With TBMargesBasse(Index)
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name Then
                ChangementMarges = True
            End If
        End If
    End With
End Sub

Private Sub TBMargesBasse_GotFocus(Index As Integer)
    On Error Resume Next
    Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub TBMargesBasse_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBMargesBasse_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_REELS_POSITIFS
End Sub

Private Sub TBMargesBasse_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    With TBMargesBasse(Index)
        If .Text = "" Then
            .Text = "0,00"
        Else
            .Text = Format(CDbl(.Text), "##0.00")
        End If
    End With
End Sub

Private Sub TBMargesDroite_Change(Index As Integer)
    On Error Resume Next
    With TBMargesDroite(Index)
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name Then
                ChangementMarges = True
            End If
        End If
    End With
End Sub

Private Sub TBMargesDroite_GotFocus(Index As Integer)
    On Error Resume Next
    Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub TBMargesDroite_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBMargesDroite_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_REELS_POSITIFS
End Sub

Private Sub TBMargesDroite_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    With TBMargesDroite(Index)
        If .Text = "" Then
            .Text = "0,00"
        Else
            .Text = Format(CDbl(.Text), "##0.00")
        End If
    End With
End Sub

Private Sub TBMargesGauche_Change(Index As Integer)
    On Error Resume Next
    With TBMargesGauche(Index)
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name Then
                ChangementMarges = True
            End If
        End If
    End With
End Sub

Private Sub TBMargesGauche_GotFocus(Index As Integer)
    On Error Resume Next
    Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub TBMargesGauche_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBMargesGauche_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_REELS_POSITIFS
End Sub

Private Sub TBMargesGauche_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    With TBMargesGauche(Index)
        If .Text = "" Then
            .Text = "0,00"
        Else
            .Text = Format(CDbl(.Text), "##0.00")
        End If
    End With
End Sub

Private Sub TBMargesHaute_Change(Index As Integer)
    On Error Resume Next
    With TBMargesHaute(Index)
        If PremiereActivation = True Then
            If Me.ActiveControl.Name = .Name Then
                ChangementMarges = True
            End If
        End If
    End With
End Sub

Private Sub TBMargesHaute_GotFocus(Index As Integer)
    On Error Resume Next
    Me.ActiveControl.SelLength = Len(Me.ActiveControl.Text)
End Sub

Private Sub TBMargesHaute_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    FiltreToucheFonction KeyCode, Shift
End Sub

Private Sub TBMargesHaute_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    FiltreToucheASCII KeyAscii, DONNEES.D_NBR_REELS_POSITIFS
End Sub

Private Sub TBMargesHaute_Validate(Index As Integer, Cancel As Boolean)
    On Error Resume Next
    With TBMargesHaute(Index)
        If .Text = "" Then
            .Text = "0,00"
        Else
            .Text = Format(CDbl(.Text), "##0.00")
        End If
    End With
End Sub

Private Sub TimerImprimanteParDefaut_Timer()

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affichage du nom de l'imprimante ---
    If LImprimanteParDefaut.Caption <> Printer.DeviceName Then
        LImprimanteParDefaut.Caption = Printer.DeviceName
    End If

    '--- bip de passage dans la routine UNIQUEMENT POUR LES TESTS ---
    If PROGRAMME_AVEC_AUTOMATE = False Then Beep

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Charge les valeurs de marges à partir de la base des registres
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ChargementMarges()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim OccOptionButton As OptionButton

    '--- lecture de l'option sélectionnée ---
    OptionSelectionnee = GetSetting(App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option sélectionnée", "1")
    If OptionSelectionnee = 0 Then OptionSelectionnee = 1
    
    '--- lecture des valeurs ---
    For Each OccOptionButton In OBOptionsImpression
        With OccOptionButton
            If .Enabled = True Then
                Me.TBMargesGauche(.Index) = GetSetting(App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge gauche", "0,00")
                Me.TBMargesHaute(.Index) = GetSetting(App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge haute", "0,00")
                Me.TBMargesDroite(.Index) = GetSetting(App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge droite", "0,00")
                Me.TBMargesBasse(.Index) = GetSetting(App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge basse", "0,00")
            End If
        End With
    Next

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Charge le nom de la personne émettrice
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub ChargementPersonneEmettrice()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- lecture dans la base de registre ---
    PersonneEmettrice = GetSetting(App.Title, Me.Name, "Personne émettrice", "")

    '--- affectation ---
    DCPersonneEmettrice.Text = PersonneEmettrice

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre le nom de la personne émettrice
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistrementPersonneEmettrice()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    PersonneEmettrice = DCPersonneEmettrice.Text
    
    '--- enregistrement dans la base de registre ---
    SaveSetting App.Title, Me.Name, "Personne émettrice", PersonneEmettrice

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistre les valeurs de marges dans la base des registres
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub EnregistrementMarges()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim OccOptionButton As OptionButton
    
    '--- enregistrement de l'option sélectionnée ---
    SaveSetting App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option sélectionnée", OptionSelectionnee
    
    '--- enregistrement ---
    If ChangementMarges = True Then
        For Each OccOptionButton In OBOptionsImpression
            With OccOptionButton
                If .Enabled = True Then
                    SaveSetting App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge gauche", Me.TBMargesGauche(.Index)
                    SaveSetting App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge haute", Me.TBMargesHaute(.Index)
                    SaveSetting App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge droite", Me.TBMargesDroite(.Index)
                    SaveSetting App.Title, Me.Name, "fenetre n° " & CStr(NumfenetreAppel) & ", option " & .Index & ", marge basse", Me.TBMargesBasse(.Index)
                End If
            End With
        Next
    End If

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le calcul des marges d'impression en twips
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CalculMargesImpressionTwips()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- affectation ---
    MargeGaucheTwips = 0
    MargeHauteTwips = 0
    MargeDroiteTwips = 0
    MargeBasseTwips = 0

    '--- calculs de marges  ---
    With TBMargesGauche(OptionSelectionnee)
        If IsNumeric(.Text) = True Then
            MargeGaucheTwips = CLng(CDbl(.Text) * 56.7)
        End If
    End With
    With TBMargesHaute(OptionSelectionnee)
        If IsNumeric(.Text) = True Then
            MargeHauteTwips = CLng(CDbl(.Text) * 56.7)
        End If
    End With
    With TBMargesDroite(OptionSelectionnee)
        If IsNumeric(.Text) = True Then
            MargeDroiteTwips = CLng(CDbl(.Text) * 56.7)
        End If
    End With
    With TBMargesBasse(OptionSelectionnee)
        If IsNumeric(.Text) = True Then
            MargeBasseTwips = CLng(CDbl(.Text) * 56.7)
        End If
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    
    '--- neutralisation du timer ---
    With TimerImprimanteParDefaut
        .Enabled = False
        .Interval = 0
    End With

    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set FChoixImpression = Nothing
    DoEvents

End Sub


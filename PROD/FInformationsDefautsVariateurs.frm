VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FInformationsDefautsVariateurs 
   Caption         =   "INFORMATIONS SUR LES DEFAUTS DES VARIATEURS"
   ClientHeight    =   10725
   ClientLeft      =   1710
   ClientTop       =   2340
   ClientWidth     =   15405
   HelpContextID   =   80
   Icon            =   "FInformationsDefautsVariateurs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10725
   ScaleWidth      =   15405
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FInformationsDefautsVariateurs.frx":014A
      ScaleHeight     =   315
      ScaleWidth      =   15345
      TabIndex        =   370
      Top             =   0
      Width           =   15405
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "MAINTENANCE GEREE"
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
         TabIndex        =   371
         Top             =   0
         Width           =   11415
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox PBListesDefautsVariateurs 
      Height          =   795
      Index           =   6
      Left            =   0
      ScaleHeight     =   735
      ScaleWidth      =   14955
      TabIndex        =   293
      TabStop         =   0   'False
      Top             =   5940
      Visible         =   0   'False
      Width           =   15015
      Begin RichTextLib.RichTextBox RichTextBox31 
         Height          =   675
         Left            =   8460
         TabIndex        =   294
         TabStop         =   0   'False
         Top             =   5880
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":24A8C
      End
      Begin RichTextLib.RichTextBox RichTextBox32 
         Height          =   675
         Left            =   5280
         TabIndex        =   295
         TabStop         =   0   'False
         Top             =   5880
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":24B32
      End
      Begin RichTextLib.RichTextBox RichTextBox33 
         Height          =   675
         Left            =   3780
         TabIndex        =   296
         Top             =   5880
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":24C29
      End
      Begin RichTextLib.RichTextBox RichTextBox34 
         Height          =   675
         Left            =   2220
         TabIndex        =   297
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":24CC4
      End
      Begin RichTextLib.RichTextBox RichTextBox35 
         Height          =   675
         Left            =   1380
         TabIndex        =   298
         TabStop         =   0   'False
         Top             =   5880
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":24D5B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox36 
         Height          =   2055
         Left            =   5280
         TabIndex        =   299
         TabStop         =   0   'False
         Top             =   3840
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   3625
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":24DDA
      End
      Begin RichTextLib.RichTextBox RichTextBox37 
         Height          =   2055
         Left            =   3780
         TabIndex        =   300
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   3625
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2505B
      End
      Begin RichTextLib.RichTextBox RichTextBox38 
         Height          =   2055
         Left            =   2220
         TabIndex        =   301
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3625
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":250EC
      End
      Begin RichTextLib.RichTextBox RichTextBox39 
         Height          =   2055
         Left            =   1380
         TabIndex        =   302
         TabStop         =   0   'False
         Top             =   3840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   3625
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2517F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   39
         Left            =   1380
         TabIndex        =   303
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":251FE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   39
         Left            =   2220
         TabIndex        =   304
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2527D
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   40
         Left            =   5280
         TabIndex        =   305
         TabStop         =   0   'False
         Top             =   360
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2530D
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   40
         Left            =   8460
         TabIndex        =   306
         TabStop         =   0   'False
         Top             =   360
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":253E7
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   39
         Left            =   3780
         TabIndex        =   307
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25579
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1095
         Index           =   42
         Left            =   1380
         TabIndex        =   308
         TabStop         =   0   'False
         Top             =   1020
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25614
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1095
         Index           =   42
         Left            =   2220
         TabIndex        =   309
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25693
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1095
         Index           =   42
         Left            =   3780
         TabIndex        =   310
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25728
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1095
         Index           =   43
         Left            =   5280
         TabIndex        =   311
         TabStop         =   0   'False
         Top             =   1020
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":257C3
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1095
         Index           =   41
         Left            =   8460
         TabIndex        =   312
         TabStop         =   0   'False
         Top             =   1020
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":258EB
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   43
         Left            =   1380
         TabIndex        =   313
         TabStop         =   0   'False
         Top             =   2100
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25991
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   43
         Left            =   2220
         TabIndex        =   314
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25A10
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   43
         Left            =   3780
         TabIndex        =   315
         TabStop         =   0   'False
         Top             =   2100
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25AA0
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   44
         Left            =   5280
         TabIndex        =   316
         TabStop         =   0   'False
         Top             =   2100
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25B3B
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   42
         Left            =   8460
         TabIndex        =   317
         TabStop         =   0   'False
         Top             =   2100
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25C2D
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1095
         Index           =   44
         Left            =   1380
         TabIndex        =   318
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25CE8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1095
         Index           =   44
         Left            =   2220
         TabIndex        =   319
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25D67
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1095
         Index           =   44
         Left            =   3780
         TabIndex        =   320
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25DF9
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1095
         Index           =   45
         Left            =   5280
         TabIndex        =   321
         TabStop         =   0   'False
         Top             =   2760
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25E8A
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1095
         Index           =   43
         Left            =   8460
         TabIndex        =   322
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":25FC8
      End
      Begin RichTextLib.RichTextBox RichTextBox40 
         Height          =   2055
         Left            =   8460
         TabIndex        =   323
         TabStop         =   0   'False
         Top             =   3840
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3625
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26091
      End
      Begin RichTextLib.RichTextBox RichTextBox41 
         Height          =   1875
         Left            =   1380
         TabIndex        =   324
         Top             =   6540
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   3307
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26314
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox42 
         Height          =   1875
         Left            =   2220
         TabIndex        =   325
         Top             =   6540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3307
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26393
      End
      Begin RichTextLib.RichTextBox RichTextBox43 
         Height          =   1875
         Left            =   3780
         TabIndex        =   326
         Top             =   6540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   3307
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2642F
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1875
         Index           =   46
         Left            =   5280
         TabIndex        =   327
         TabStop         =   0   'False
         Top             =   6540
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   3307
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":264C0
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1875
         Index           =   44
         Left            =   8460
         TabIndex        =   328
         TabStop         =   0   'False
         Top             =   6540
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   3307
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":267B3
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remède"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   5
         Left            =   8460
         TabIndex        =   366
         Top             =   60
         Width           =   5295
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cause possible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   5
         Left            =   5280
         TabIndex        =   365
         Top             =   60
         Width           =   3195
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   5
         Left            =   4980
         TabIndex        =   364
         Top             =   60
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Réaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   5
         Left            =   3780
         TabIndex        =   363
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Désignation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   5
         Left            =   2220
         TabIndex        =   362
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Défaut n°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   5
         Left            =   1380
         TabIndex        =   361
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   4980
         TabIndex        =   335
         Top             =   5880
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   46
         Left            =   4980
         TabIndex        =   334
         Top             =   360
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   45
         Left            =   4980
         TabIndex        =   333
         Top             =   1020
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   44
         Left            =   4980
         TabIndex        =   332
         Top             =   2100
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   43
         Left            =   4980
         TabIndex        =   331
         Top             =   2760
         Width           =   315
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   4980
         TabIndex        =   330
         Top             =   3840
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   42
         Left            =   4980
         TabIndex        =   329
         Top             =   6540
         Width           =   315
      End
   End
   Begin VB.PictureBox PBListesDefautsVariateurs 
      Height          =   915
      Index           =   3
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   14955
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   15015
      Begin RichTextLib.RichTextBox RichTextBox5 
         Height          =   1035
         Left            =   5280
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   7380
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2685E
      End
      Begin RichTextLib.RichTextBox RichTextBox4 
         Height          =   1035
         Left            =   3780
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   7380
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":269B5
      End
      Begin RichTextLib.RichTextBox RichTextBox3 
         Height          =   1035
         Left            =   2220
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   7380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26A50
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   1035
         Left            =   1380
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   7380
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26AE7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   435
         Left            =   7980
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   6960
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   767
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26B66
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   435
         Index           =   28
         Left            =   5280
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   6960
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   767
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26C0C
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   19
         Left            =   1380
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26C9F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   19
         Left            =   2220
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26D1E
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   19
         Left            =   5280
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   360
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26DB2
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   19
         Left            =   7980
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   360
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26E6B
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   19
         Left            =   3780
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":26F87
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   20
         Left            =   1380
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27018
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   20
         Left            =   2220
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27097
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   20
         Left            =   3780
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27131
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   20
         Left            =   5280
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   840
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":271CC
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   20
         Left            =   7980
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   840
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":272A2
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1455
         Index           =   21
         Left            =   1380
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   1320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":273CC
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1455
         Index           =   21
         Left            =   2220
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2744B
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1455
         Index           =   21
         Left            =   3780
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2566
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":274E2
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1455
         Index           =   21
         Left            =   5280
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27573
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1455
         Index           =   21
         Left            =   7980
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1320
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":276EA
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   23
         Left            =   1380
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   3600
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":278FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   23
         Left            =   2220
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27979
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   23
         Left            =   3780
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27A0A
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   23
         Left            =   5280
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27AA5
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   23
         Left            =   7980
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   3600
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27B53
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   24
         Left            =   1380
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   4080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27BFE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   24
         Left            =   2220
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27C7D
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   24
         Left            =   3780
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27D0B
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   24
         Left            =   5280
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   4080
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27DA6
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   24
         Left            =   7980
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   4080
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27E86
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   25
         Left            =   1380
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   4740
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":27F99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   25
         Left            =   2220
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28018
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   25
         Left            =   3780
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":280AA
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   25
         Left            =   5280
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   4740
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28145
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   25
         Left            =   7980
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   4740
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":281F2
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1275
         Index           =   26
         Left            =   1380
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   5220
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2249
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":282AD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1275
         Index           =   26
         Left            =   2220
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   5220
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2249
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2832C
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1275
         Index           =   26
         Left            =   3780
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   5220
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2249
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":283BB
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1275
         Index           =   26
         Left            =   5280
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   5220
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2249
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28456
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1275
         Index           =   26
         Left            =   7980
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   5220
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2249
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":285AA
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   27
         Left            =   1380
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28825
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   27
         Left            =   2220
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":288A4
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   27
         Left            =   3780
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28936
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   27
         Left            =   5280
         TabIndex        =   170
         TabStop         =   0   'False
         Top             =   6480
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":289D1
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   27
         Left            =   7980
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   6480
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28A7F
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   435
         Index           =   28
         Left            =   1380
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   6960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   767
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28B25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   435
         Index           =   28
         Left            =   2220
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28BA4
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   435
         Index           =   28
         Left            =   3780
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   767
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28C36
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   22
         Left            =   1380
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28CD1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   22
         Left            =   2220
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28D50
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   22
         Left            =   3780
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28DE7
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   22
         Left            =   5280
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28E78
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   22
         Left            =   7980
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":28F5B
      End
      Begin RichTextLib.RichTextBox RichTextBox6 
         Height          =   1035
         Left            =   7980
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   7380
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":29033
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Défaut n°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   2
         Left            =   1380
         TabIndex        =   348
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Désignation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   2
         Left            =   2220
         TabIndex        =   347
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Réaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   2
         Left            =   3780
         TabIndex        =   346
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   2
         Left            =   4980
         TabIndex        =   345
         Top             =   60
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cause possible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   2
         Left            =   5280
         TabIndex        =   344
         Top             =   60
         Width           =   2715
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remède"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   2
         Left            =   7980
         TabIndex        =   343
         Top             =   60
         Width           =   5715
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   4980
         TabIndex        =   181
         Top             =   7380
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   28
         Left            =   4980
         TabIndex        =   175
         Top             =   6960
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   27
         Left            =   4980
         TabIndex        =   169
         Top             =   6480
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1275
         Index           =   26
         Left            =   4980
         TabIndex        =   163
         Top             =   5220
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   25
         Left            =   4980
         TabIndex        =   157
         Top             =   4740
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   24
         Left            =   4980
         TabIndex        =   151
         Top             =   4080
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   23
         Left            =   4980
         TabIndex        =   150
         Top             =   3600
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   22
         Left            =   4980
         TabIndex        =   141
         Top             =   2760
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   21
         Left            =   4980
         TabIndex        =   138
         Top             =   1320
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   20
         Left            =   4980
         TabIndex        =   132
         Top             =   840
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   19
         Left            =   4980
         TabIndex        =   128
         Top             =   360
         Width           =   315
      End
   End
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   15345
      TabIndex        =   336
      TabStop         =   0   'False
      Top             =   9630
      Width           =   15405
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FInformationsDefautsVariateurs.frx":29219
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
         Left            =   13740
         MaskColor       =   &H00FF00FF&
         Picture         =   "FInformationsDefautsVariateurs.frx":2991B
         Style           =   1  'Graphical
         TabIndex        =   369
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.CommandButton CBPrecedent 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Précédent"
         DownPicture     =   "FInformationsDefautsVariateurs.frx":2A01D
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
         Left            =   10260
         MaskColor       =   &H00FF00FF&
         Picture         =   "FInformationsDefautsVariateurs.frx":2A71F
         Style           =   1  'Graphical
         TabIndex        =   367
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.CommandButton CBSuivant 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Suivant"
         DownPicture     =   "FInformationsDefautsVariateurs.frx":2AE21
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
         Left            =   12000
         MaskColor       =   &H00FF00FF&
         Picture         =   "FInformationsDefautsVariateurs.frx":2B523
         Style           =   1  'Graphical
         TabIndex        =   368
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1575
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   2460
         Top             =   240
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin VB.PictureBox PBListesDefautsVariateurs 
      Height          =   1215
      Index           =   1
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   14955
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   15015
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   1
         Left            =   7980
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   960
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2BC25
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   1
         Left            =   3780
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2BDCE
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   1
         Left            =   1380
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2BE69
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   1
         Left            =   2220
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2BEE8
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   1
         Left            =   5280
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   960
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2BF76
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   255
         Index           =   0
         Left            =   1380
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C07C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   255
         Index           =   0
         Left            =   2220
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C0FB
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   3
         Left            =   1380
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1620
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C18A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   2
         Left            =   2220
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C209
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   0
         Left            =   3780
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C2AE
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   0
         Left            =   5280
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1620
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C349
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   0
         Left            =   7980
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1620
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C46C
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1455
         Index           =   4
         Left            =   1380
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2460
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C589
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1455
         Index           =   3
         Left            =   2220
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C608
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1455
         Index           =   2
         Left            =   3780
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2566
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C694
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1455
         Index           =   2
         Left            =   5280
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2460
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C72F
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1455
         Index           =   2
         Left            =   7980
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2460
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2566
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2C8DB
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   6
         Left            =   1380
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3900
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CADF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   4
         Left            =   2220
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3900
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CB5E
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   3
         Left            =   3780
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CBF7
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   3
         Left            =   5280
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   3900
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CC92
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   3
         Left            =   7980
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3900
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CD23
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   2
         Left            =   1380
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   4380
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CDC8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   5
         Left            =   2220
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4380
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CE47
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   4
         Left            =   3780
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   4380
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CEE9
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   4
         Left            =   5280
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   4380
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2CF84
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   4
         Left            =   7980
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4380
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D03A
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1875
         Index           =   5
         Left            =   1380
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   5040
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   3307
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D1E7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1875
         Index           =   6
         Left            =   2220
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3307
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D266
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1875
         Index           =   5
         Left            =   3780
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   3307
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D2F2
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1875
         Index           =   5
         Left            =   5280
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   5040
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   3307
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D38D
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1875
         Index           =   5
         Left            =   7980
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   5040
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3307
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D562
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1035
         Index           =   7
         Left            =   1380
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   6900
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D97B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1035
         Index           =   7
         Left            =   2220
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   6900
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2D9FA
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1035
         Index           =   6
         Left            =   3780
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   6900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DA88
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1035
         Index           =   6
         Left            =   5280
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6900
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1826
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DB23
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1035
         Index           =   6
         Left            =   7980
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   6900
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1826
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DC0F
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   255
         Index           =   18
         Left            =   3780
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DCD6
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   255
         Index           =   18
         Left            =   7980
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   720
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DD58
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   255
         Index           =   18
         Left            =   5280
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   720
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   450
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DDDA
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1035
         Index           =   18
         Left            =   4980
         TabIndex        =   119
         Top             =   6900
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   17
         Left            =   4980
         TabIndex        =   118
         Top             =   5040
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   16
         Left            =   4980
         TabIndex        =   117
         Top             =   4380
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   15
         Left            =   4980
         TabIndex        =   116
         Top             =   3900
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1455
         Index           =   14
         Left            =   4980
         TabIndex        =   115
         Top             =   2460
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   13
         Left            =   4980
         TabIndex        =   114
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   12
         Left            =   4980
         TabIndex        =   113
         Top             =   960
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   4980
         TabIndex        =   112
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Défaut n°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   0
         Left            =   1380
         TabIndex        =   11
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Désignation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   0
         Left            =   2220
         TabIndex        =   10
         Top             =   420
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Réaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   0
         Left            =   3780
         TabIndex        =   9
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   0
         Left            =   4980
         TabIndex        =   8
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cause possible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   0
         Left            =   5280
         TabIndex        =   7
         Top             =   420
         Width           =   2715
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remède"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   0
         Left            =   7980
         TabIndex        =   6
         Top             =   420
         Width           =   5715
      End
   End
   Begin VB.PictureBox PBListesDefautsVariateurs 
      Height          =   1215
      Index           =   2
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   14955
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   15015
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1095
         Index           =   8
         Left            =   1380
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   540
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DE5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1095
         Index           =   8
         Left            =   2220
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DEDB
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1095
         Index           =   7
         Left            =   3780
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1931
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DF64
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1095
         Index           =   7
         Left            =   5280
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   540
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2DFF5
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1095
         Index           =   7
         Left            =   7980
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   540
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1931
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E12C
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   315
         Index           =   9
         Left            =   1380
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1620
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E2B4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   315
         Index           =   9
         Left            =   2220
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E333
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   315
         Index           =   8
         Left            =   3780
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E3C3
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   315
         Index           =   8
         Left            =   5280
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1620
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E454
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   315
         Index           =   8
         Left            =   7980
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1620
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   556
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E4F3
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   10
         Left            =   1380
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E5AF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   10
         Left            =   2220
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E62E
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   9
         Left            =   3780
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E6BD
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   9
         Left            =   5280
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E74E
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   9
         Left            =   7980
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1920
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E845
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   11
         Left            =   1380
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2760
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E93F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   11
         Left            =   2220
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2E9BE
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   10
         Left            =   3780
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2EA49
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   10
         Left            =   5280
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2EAE4
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   10
         Left            =   7980
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2EB92
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1035
         Index           =   12
         Left            =   1380
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2EC3D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1035
         Index           =   12
         Left            =   2220
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2ECBC
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1035
         Index           =   11
         Left            =   3780
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1826
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2ED41
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1035
         Index           =   11
         Left            =   5280
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   3240
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1826
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2EDDC
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1035
         Index           =   11
         Left            =   7980
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3240
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1826
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2EF25
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   13
         Left            =   1380
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4260
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F018
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   13
         Left            =   2220
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   4260
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F097
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   12
         Left            =   3780
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   4260
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F122
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   12
         Left            =   5280
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   4260
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F1BD
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   12
         Left            =   7980
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   4260
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F25F
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   14
         Left            =   1380
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   4740
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F34A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   14
         Left            =   2220
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F3C9
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   13
         Left            =   3780
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F45C
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   13
         Left            =   5280
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   4740
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F4F7
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   13
         Left            =   7980
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4740
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F5E1
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   615
         Index           =   15
         Left            =   1380
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1085
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F6F7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   615
         Index           =   15
         Left            =   2220
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F776
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   615
         Index           =   14
         Left            =   3780
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1085
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F7FB
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   615
         Index           =   14
         Left            =   5280
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   5400
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1085
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F889
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   615
         Index           =   14
         Left            =   7980
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   5400
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1085
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2F942
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   16
         Left            =   1380
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   6000
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FAAB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   16
         Left            =   2220
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FB2A
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   15
         Left            =   3780
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FBB6
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   15
         Left            =   5280
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   6000
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FC47
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   15
         Left            =   7980
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   6000
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FD34
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   17
         Left            =   1380
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   6660
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FDFA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   17
         Left            =   2220
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   6660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FE79
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   16
         Left            =   3780
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   6660
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FF0F
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   16
         Left            =   5280
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   6660
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":2FFA0
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   16
         Left            =   7980
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   6660
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":300B0
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   18
         Left            =   1380
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   7500
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3022C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   18
         Left            =   2220
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   7500
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":302AB
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   17
         Left            =   5280
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   7500
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30341
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   17
         Left            =   7980
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   7500
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3041C
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   17
         Left            =   3780
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   7500
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30558
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remède"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   1
         Left            =   7980
         TabIndex        =   342
         Top             =   240
         Width           =   5715
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cause possible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   1
         Left            =   5280
         TabIndex        =   341
         Top             =   240
         Width           =   2715
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   1
         Left            =   4980
         TabIndex        =   340
         Top             =   240
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Réaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   1
         Left            =   3780
         TabIndex        =   339
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Désignation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   1
         Left            =   2220
         TabIndex        =   338
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Défaut n°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   1
         Left            =   1380
         TabIndex        =   337
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   10
         Left            =   4980
         TabIndex        =   111
         Top             =   6660
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Index           =   9
         Left            =   4980
         TabIndex        =   110
         Top             =   5400
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   8
         Left            =   4980
         TabIndex        =   109
         Top             =   4740
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   4980
         TabIndex        =   108
         Top             =   4260
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1035
         Index           =   6
         Left            =   4980
         TabIndex        =   107
         Top             =   3240
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   4980
         TabIndex        =   106
         Top             =   2760
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   4980
         TabIndex        =   105
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   4980
         TabIndex        =   104
         Top             =   1620
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   2
         Left            =   4980
         TabIndex        =   103
         Top             =   540
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   1
         Left            =   4980
         TabIndex        =   102
         Top             =   7500
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   0
         Left            =   4980
         TabIndex        =   101
         Top             =   6000
         Width           =   315
      End
   End
   Begin VB.PictureBox PBListesDefautsVariateurs 
      Height          =   855
      Index           =   4
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14955
      TabIndex        =   189
      TabStop         =   0   'False
      Top             =   4020
      Visible         =   0   'False
      Width           =   15015
      Begin RichTextLib.RichTextBox RichTextBox17 
         Height          =   1035
         Left            =   7980
         TabIndex        =   249
         TabStop         =   0   'False
         Top             =   7320
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1826
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":305E9
      End
      Begin RichTextLib.RichTextBox RichTextBox16 
         Height          =   1035
         Left            =   5280
         TabIndex        =   248
         TabStop         =   0   'False
         Top             =   7320
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1826
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30763
      End
      Begin RichTextLib.RichTextBox RichTextBox15 
         Height          =   1035
         Left            =   3780
         TabIndex        =   246
         TabStop         =   0   'False
         Top             =   7320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1826
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":308C5
      End
      Begin RichTextLib.RichTextBox RichTextBox14 
         Height          =   1035
         Left            =   2220
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   7320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1826
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30956
      End
      Begin RichTextLib.RichTextBox RichTextBox13 
         Height          =   1035
         Left            =   1380
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   7320
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1826
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":309F5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox7 
         Height          =   855
         Left            =   5280
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   6480
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30A74
      End
      Begin RichTextLib.RichTextBox RichTextBox8 
         Height          =   855
         Left            =   3780
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30B99
      End
      Begin RichTextLib.RichTextBox RichTextBox9 
         Height          =   855
         Left            =   2220
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30C34
      End
      Begin RichTextLib.RichTextBox RichTextBox10 
         Height          =   855
         Left            =   1380
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   6480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30CBF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox11 
         Height          =   495
         Left            =   7980
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   6000
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30D3E
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   29
         Left            =   5280
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   6000
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30DEA
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   29
         Left            =   1380
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30EA3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   29
         Left            =   2220
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30F22
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   30
         Left            =   5280
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   420
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":30FB4
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   28
         Left            =   7980
         TabIndex        =   199
         TabStop         =   0   'False
         Top             =   420
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":310C9
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   29
         Left            =   3780
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3117B
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   30
         Left            =   1380
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   1260
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   30
         Left            =   2220
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31295
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   30
         Left            =   3780
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31323
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   31
         Left            =   5280
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   1260
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":313BE
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   29
         Left            =   7980
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31492
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1635
         Index           =   31
         Left            =   1380
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   1920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2884
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31538
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1635
         Index           =   31
         Left            =   2220
         TabIndex        =   207
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2884
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":315B7
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1635
         Index           =   31
         Left            =   3780
         TabIndex        =   208
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2884
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31649
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1635
         Index           =   32
         Left            =   5280
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2884
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":316E4
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1635
         Index           =   30
         Left            =   7980
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   1920
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2884
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":318AD
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   32
         Left            =   1380
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   3540
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31C2D
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   32
         Left            =   2220
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31CAC
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   32
         Left            =   3780
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31D3A
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   33
         Left            =   5280
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   3540
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31DC8
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   31
         Left            =   7980
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   3540
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31E76
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   675
         Index           =   33
         Left            =   1380
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   4020
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31F5C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   675
         Index           =   33
         Left            =   2220
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   4020
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":31FDB
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   675
         Index           =   33
         Left            =   3780
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   4020
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3206D
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   675
         Index           =   34
         Left            =   5280
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   4020
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32108
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   675
         Index           =   32
         Left            =   7980
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   4020
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":321C0
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   35
         Left            =   1380
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   4680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32381
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   35
         Left            =   2220
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32400
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   35
         Left            =   3780
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3248E
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   36
         Left            =   5280
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32529
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   34
         Left            =   7980
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   4680
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32640
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   36
         Left            =   1380
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   5520
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":327B8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   36
         Left            =   2220
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32837
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   36
         Left            =   3780
         TabIndex        =   228
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":328CF
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   37
         Left            =   5280
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   5520
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3295D
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   35
         Left            =   7980
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   5520
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32A16
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   37
         Left            =   1380
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   6000
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32AC2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   37
         Left            =   2220
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32B41
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   37
         Left            =   3780
         TabIndex        =   233
         TabStop         =   0   'False
         Top             =   6000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32BD9
      End
      Begin RichTextLib.RichTextBox RichTextBox12 
         Height          =   855
         Left            =   7980
         TabIndex        =   234
         TabStop         =   0   'False
         Top             =   6480
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32C74
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remède"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   3
         Left            =   7980
         TabIndex        =   354
         Top             =   120
         Width           =   5715
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cause possible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   3
         Left            =   5280
         TabIndex        =   353
         Top             =   120
         Width           =   2715
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   3
         Left            =   4980
         TabIndex        =   352
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Réaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   3
         Left            =   3780
         TabIndex        =   351
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Désignation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   3
         Left            =   2220
         TabIndex        =   350
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Défaut n°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   3
         Left            =   1380
         TabIndex        =   349
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   4980
         TabIndex        =   247
         Top             =   7320
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   38
         Left            =   4980
         TabIndex        =   243
         Top             =   420
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   37
         Left            =   4980
         TabIndex        =   242
         Top             =   1260
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "  ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1635
         Index           =   36
         Left            =   4980
         TabIndex        =   241
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   34
         Left            =   4980
         TabIndex        =   240
         Top             =   3540
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Index           =   33
         Left            =   4980
         TabIndex        =   239
         Top             =   4020
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   31
         Left            =   4980
         TabIndex        =   238
         Top             =   4680
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   30
         Left            =   4980
         TabIndex        =   237
         Top             =   5520
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   29
         Left            =   4980
         TabIndex        =   236
         Top             =   6000
         Width           =   315
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   4980
         TabIndex        =   235
         Top             =   6480
         Width           =   315
      End
   End
   Begin VB.PictureBox PBListesDefautsVariateurs 
      Height          =   855
      Index           =   5
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   14955
      TabIndex        =   250
      TabStop         =   0   'False
      Top             =   4980
      Visible         =   0   'False
      Width           =   15015
      Begin RichTextLib.RichTextBox RichTextBox18 
         Height          =   675
         Left            =   7980
         TabIndex        =   251
         TabStop         =   0   'False
         Top             =   7260
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32D8E
      End
      Begin RichTextLib.RichTextBox RichTextBox19 
         Height          =   675
         Left            =   5280
         TabIndex        =   252
         TabStop         =   0   'False
         Top             =   7260
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32F30
      End
      Begin RichTextLib.RichTextBox RichTextBox20 
         Height          =   675
         Left            =   3780
         TabIndex        =   253
         TabStop         =   0   'False
         Top             =   7260
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":32FDD
      End
      Begin RichTextLib.RichTextBox RichTextBox21 
         Height          =   675
         Left            =   2220
         TabIndex        =   254
         TabStop         =   0   'False
         Top             =   7260
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3306E
      End
      Begin RichTextLib.RichTextBox RichTextBox22 
         Height          =   675
         Left            =   1380
         TabIndex        =   255
         TabStop         =   0   'False
         Top             =   7260
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1191
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33108
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox23 
         Height          =   1455
         Left            =   5280
         TabIndex        =   256
         Top             =   5820
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   2566
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33187
      End
      Begin RichTextLib.RichTextBox RichTextBox24 
         Height          =   1455
         Left            =   3780
         TabIndex        =   257
         Top             =   5820
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   2566
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33308
      End
      Begin RichTextLib.RichTextBox RichTextBox25 
         Height          =   1455
         Left            =   2220
         TabIndex        =   258
         Top             =   5820
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   2566
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":333A3
      End
      Begin RichTextLib.RichTextBox RichTextBox26 
         Height          =   1455
         Left            =   1380
         TabIndex        =   259
         Top             =   5820
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   2566
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33430
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   855
         Index           =   34
         Left            =   1380
         TabIndex        =   260
         TabStop         =   0   'False
         Top             =   420
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":334AF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   855
         Index           =   34
         Left            =   2220
         TabIndex        =   261
         TabStop         =   0   'False
         Top             =   420
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3352E
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   855
         Index           =   38
         Left            =   5280
         TabIndex        =   262
         TabStop         =   0   'False
         Top             =   420
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":335CD
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   855
         Index           =   33
         Left            =   7980
         TabIndex        =   263
         TabStop         =   0   'False
         Top             =   420
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   1508
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":336FD
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   855
         Index           =   34
         Left            =   3780
         TabIndex        =   264
         TabStop         =   0   'False
         Top             =   420
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1508
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33848
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   2235
         Index           =   38
         Left            =   1380
         TabIndex        =   265
         TabStop         =   0   'False
         Top             =   1260
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   3942
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":338D9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   2235
         Index           =   38
         Left            =   2220
         TabIndex        =   266
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3942
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33958
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   2235
         Index           =   38
         Left            =   3780
         TabIndex        =   267
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   3942
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33A07
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   2235
         Index           =   39
         Left            =   5280
         TabIndex        =   268
         TabStop         =   0   'False
         Top             =   1260
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   3942
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33AA2
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   2235
         Index           =   36
         Left            =   7980
         TabIndex        =   269
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3942
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33CF9
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   495
         Index           =   40
         Left            =   1380
         TabIndex        =   270
         TabStop         =   0   'False
         Top             =   3480
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33DF8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   495
         Index           =   40
         Left            =   2220
         TabIndex        =   271
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33E77
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   495
         Index           =   40
         Left            =   3780
         TabIndex        =   272
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33EFE
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   41
         Left            =   5280
         TabIndex        =   273
         TabStop         =   0   'False
         Top             =   3480
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":33F99
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   38
         Left            =   7980
         TabIndex        =   274
         TabStop         =   0   'False
         Top             =   3480
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3404A
      End
      Begin RichTextLib.RichTextBox RTBNumDefaut 
         Height          =   1875
         Index           =   41
         Left            =   1380
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   3960
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   3307
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":340F0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTBDesignation 
         Height          =   1875
         Index           =   41
         Left            =   2220
         TabIndex        =   276
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   3307
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":3416F
      End
      Begin RichTextLib.RichTextBox RTBReaction 
         Height          =   1875
         Index           =   41
         Left            =   3780
         TabIndex        =   277
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   3307
         _Version        =   393217
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":34208
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   1875
         Index           =   42
         Left            =   5280
         TabIndex        =   278
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   3307
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":342A3
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   1875
         Index           =   39
         Left            =   7980
         TabIndex        =   279
         TabStop         =   0   'False
         Top             =   3960
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   3307
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":34486
      End
      Begin RichTextLib.RichTextBox RichTextBox28 
         Height          =   1455
         Left            =   7980
         TabIndex        =   280
         Top             =   5820
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   2566
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":346DB
      End
      Begin RichTextLib.RichTextBox RichTextBox27 
         Height          =   495
         Left            =   1380
         TabIndex        =   287
         Top             =   7920
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":34854
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox RichTextBox29 
         Height          =   495
         Left            =   2220
         TabIndex        =   288
         Top             =   7920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":348D3
      End
      Begin RichTextLib.RichTextBox RichTextBox30 
         Height          =   495
         Left            =   3780
         TabIndex        =   289
         Top             =   7920
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         _Version        =   393217
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":34959
      End
      Begin RichTextLib.RichTextBox RTBCausePossible 
         Height          =   495
         Index           =   35
         Left            =   5280
         TabIndex        =   291
         Top             =   7920
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":349F4
      End
      Begin RichTextLib.RichTextBox RTBRemede 
         Height          =   495
         Index           =   37
         Left            =   7980
         TabIndex        =   292
         Top             =   7920
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   873
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         Appearance      =   0
         TextRTF         =   $"FInformationsDefautsVariateurs.frx":34A9E
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remède"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   4
         Left            =   7980
         TabIndex        =   360
         Top             =   120
         Width           =   5715
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cause possible"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   4
         Left            =   5280
         TabIndex        =   359
         Top             =   120
         Width           =   2715
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   4
         Left            =   4980
         TabIndex        =   358
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Réaction"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   4
         Left            =   3780
         TabIndex        =   357
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Désignation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   4
         Left            =   2220
         TabIndex        =   356
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Défaut n°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000014&
         Height          =   315
         Index           =   4
         Left            =   1380
         TabIndex        =   355
         Top             =   120
         Width           =   855
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   39
         Left            =   4980
         TabIndex        =   290
         Top             =   7920
         Width           =   315
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   4980
         TabIndex        =   286
         Top             =   5820
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1875
         Index           =   41
         Left            =   4980
         TabIndex        =   285
         Top             =   3960
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   40
         Left            =   4980
         TabIndex        =   284
         Top             =   3480
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2235
         Index           =   35
         Left            =   4980
         TabIndex        =   283
         Top             =   1260
         Width           =   315
      End
      Begin VB.Label LP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   32
         Left            =   4980
         TabIndex        =   282
         Top             =   420
         Width           =   315
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   4980
         TabIndex        =   281
         Top             =   7260
         Width           =   315
      End
   End
End
Attribute VB_Name = "FInformationsDefautsVariateurs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre affichant les informations sur les défauts des variateurs
' Nom                    : FInformationsDefautsVariateurs.frm
' Date de création : 12/09/2005
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z
    
'--- constantes privées ---
Private Const TITRE_FENETRE As String = "Informations sur les défauts des variateurs"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---

'--- types privées ---
    
'--- variables privées ---
Private PremiereActivation As Boolean
Private IdxImageChoisie As Integer                     'index de l'image choisie pour l'affichage

'--- variables publiques ---
Public NumFenetre As Long                                   'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBPrecedent_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- passer à l'image suivante / contrôle des limites ---
    Dec IdxImageChoisie
    If IdxImageChoisie < PBListesDefautsVariateurs.LBound Then
        IdxImageChoisie = PBListesDefautsVariateurs.LBound
    End If

    '--- affichage de l'image ---
    AfficheImage IdxImageChoisie
        
End Sub

Private Sub CBPrecedent_GotFocus()
    
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

Private Sub CBPrecedent_LostFocus()
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

Private Sub CBSuivant_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- passer à l'image suivante / contrôle des limites ---
    Inc IdxImageChoisie
    If IdxImageChoisie > PBListesDefautsVariateurs.UBound Then
        IdxImageChoisie = PBListesDefautsVariateurs.UBound
    End If

    '--- affichage de l'image ---
    AfficheImage IdxImageChoisie

End Sub

Private Sub CBSuivant_GotFocus()
    
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

Private Sub CBSuivant_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
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
    Dim a As Integer
    
    '--- agrandir les images à la taille de la Fenetre ---
    For a = PBListesDefautsVariateurs.LBound To PBListesDefautsVariateurs.UBound
        With PBListesDefautsVariateurs(a)
            .Width = Me.ScaleWidth
            .Height = Me.ScaleHeight - PBBoutons.Height - PBRenseignementsFenetre.Height
            .Top = PBRenseignementsFenetre.Height
            .Left = 0
            Set .Picture = ImgFondDeFenetre
        End With
    Next

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
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBSuivant.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBSuivant.Width
    CBPrecedent.Left = CBSuivant.Left - MARGES.M_ENTRE_BOUTONS - CBPrecedent.Width
    
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
' Rôle      : Initialise la fenêtre (chargement ou en vue de la rendre visible)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub InitialisationFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---

    '--- affectation ---
  
    '--- divers sur la Fenetre ---
    With Me
        .Caption = UCase(TITRE_FENETRE)
        .WindowState = vbMaximized
    End With
    
    '--- renseignements de la fenêtre ---
    LRenseignementsFenetre.Caption = UCase(TITRE_FENETRE)
    
    '--- fond de l'image des boutons ---
    PBBoutons.Picture = ImgFondDesBoutons
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue le paramètrage de la Fenetre
' Entrées :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ParametrageFenetre()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- forcer l'index à la première image ---
    IdxImageChoisie = 1

    '--- affichage de l'image ---
    AfficheImage IdxImageChoisie

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affiche l'image en fonction de la sélection
' Entrées : IdxImage -> Index de l'image sélectionnée
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub AfficheImage(ByVal IdxImage As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- rendre visible ou invisible les images en fonction de l'index ---
    Select Case IdxImage
        
        Case 1
            PBListesDefautsVariateurs(1).Visible = True
            PBListesDefautsVariateurs(2).Visible = False
            PBListesDefautsVariateurs(3).Visible = False
            PBListesDefautsVariateurs(4).Visible = False
            PBListesDefautsVariateurs(5).Visible = False
            PBListesDefautsVariateurs(6).Visible = False
        
        Case 2
            PBListesDefautsVariateurs(1).Visible = False
            PBListesDefautsVariateurs(2).Visible = True
            PBListesDefautsVariateurs(3).Visible = False
            PBListesDefautsVariateurs(4).Visible = False
            PBListesDefautsVariateurs(5).Visible = False
            PBListesDefautsVariateurs(6).Visible = False
        
        Case 3
            PBListesDefautsVariateurs(1).Visible = False
            PBListesDefautsVariateurs(2).Visible = False
            PBListesDefautsVariateurs(3).Visible = True
            PBListesDefautsVariateurs(4).Visible = False
            PBListesDefautsVariateurs(5).Visible = False
            PBListesDefautsVariateurs(6).Visible = False
        
        Case 4
            PBListesDefautsVariateurs(1).Visible = False
            PBListesDefautsVariateurs(2).Visible = False
            PBListesDefautsVariateurs(3).Visible = False
            PBListesDefautsVariateurs(4).Visible = True
            PBListesDefautsVariateurs(5).Visible = False
            PBListesDefautsVariateurs(6).Visible = False
        
        Case 5
            PBListesDefautsVariateurs(1).Visible = False
            PBListesDefautsVariateurs(2).Visible = False
            PBListesDefautsVariateurs(3).Visible = False
            PBListesDefautsVariateurs(4).Visible = False
            PBListesDefautsVariateurs(5).Visible = True
            PBListesDefautsVariateurs(6).Visible = False
        
        Case 6
            PBListesDefautsVariateurs(1).Visible = False
            PBListesDefautsVariateurs(2).Visible = False
            PBListesDefautsVariateurs(3).Visible = False
            PBListesDefautsVariateurs(4).Visible = False
            PBListesDefautsVariateurs(5).Visible = False
            PBListesDefautsVariateurs(6).Visible = True
        
        Case Else
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la Fenetre
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
    Set OccFInformationsDefautsVariateurs = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change le curseur de la souris en fonction de l'attente
' Entrées : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par défaut
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub SourisEnAttente(ByVal AttenteOuiNon As Boolean)
    
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
' Rôle      : Gère l'appui des touches du clavier
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub GestionTouches(KeyCode As Integer, Shift As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- action en fonction des touches ---
    Select Case KeyCode
        
        Case vbKeyPageUp
            '--- touche page en haut ---
            CBPrecedent_Click
        
        Case vbKeyPageDown
            '--- touche page en bas ---
            CBSuivant_Click
        
        Case Else
    End Select

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


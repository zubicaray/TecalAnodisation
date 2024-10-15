VERSION 5.00
Object = "{0842D103-1E19-101B-9AAF-1A1626551E7C}#1.0#0"; "GRAPH32.OCX"
Begin VB.Form FVisualisationGraphesProduction 
   Caption         =   "Visualisation des graphes de production"
   ClientHeight    =   10290
   ClientLeft      =   1695
   ClientTop       =   2460
   ClientWidth     =   17055
   Icon            =   "FVisualisationGraphesProduction.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   10290
   ScaleWidth      =   17055
   WindowState     =   2  'Maximized
   Begin VB.Frame FChoixMesures 
      Caption         =   " Choix des mesures "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9855
      Left            =   60
      TabIndex        =   6
      Top             =   360
      Width           =   4575
      Begin VB.CheckBox CBEchellesDilatees 
         Caption         =   "Echelles dilatées de U et I "
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
         Left            =   660
         TabIndex        =   12
         Top             =   1140
         Width           =   3675
      End
      Begin VB.OptionButton OBChoixGraphe 
         Caption         =   "Température de la phase 4"
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
         Index           =   3
         Left            =   165
         TabIndex        =   11
         Top             =   2175
         Width           =   4215
      End
      Begin VB.OptionButton OBChoixGraphe 
         Caption         =   "Température des phases 1 à 3"
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
         Index           =   2
         Left            =   165
         TabIndex        =   10
         Top             =   1815
         Width           =   4215
      End
      Begin VB.OptionButton OBChoixGraphe 
         Caption         =   "Tension et intensité des phases 1 à 3"
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
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   4215
      End
      Begin VB.OptionButton OBChoixGraphe 
         Caption         =   "Tension et intensité de la phase 4"
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
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   720
         Width           =   4215
      End
      Begin VB.Line LDecoration 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   4425
         X2              =   105
         Y1              =   2595
         Y2              =   2595
      End
      Begin VB.Line LDecoration 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   4425
         X2              =   105
         Y1              =   1635
         Y2              =   1635
      End
      Begin VB.Image IPhasesAnodisation 
         Height          =   2010
         Left            =   705
         Picture         =   "FVisualisationGraphesProduction.frx":0442
         Top             =   3195
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
         Left            =   705
         TabIndex        =   9
         Top             =   2895
         Width           =   2910
      End
   End
   Begin VB.PictureBox PBRenseignementsFenetre 
      Align           =   1  'Align Top
      BackColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      Picture         =   "FVisualisationGraphesProduction.frx":1384C
      ScaleHeight     =   315
      ScaleWidth      =   16995
      TabIndex        =   4
      Top             =   0
      Width           =   17055
      Begin VB.Label LRenseignementsFenetre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         BackStyle       =   0  'Transparent
         Caption         =   "VISUALISATION DES GRAPHES DE PRODUCTION"
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
         TabIndex        =   5
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
      ScaleWidth      =   16995
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   9195
      Width           =   17055
      Begin VB.CommandButton CBImprimerGraphe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Imprimer le graphe"
         DownPicture     =   "FVisualisationGraphesProduction.frx":3818E
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
         Left            =   10800
         MaskColor       =   &H00FF00FF&
         Picture         =   "FVisualisationGraphesProduction.frx":38890
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FVisualisationGraphesProduction.frx":38F92
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
         Left            =   13260
         MaskColor       =   &H00FF00FF&
         Picture         =   "FVisualisationGraphesProduction.frx":39694
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1515
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   405
         Left            =   8340
         Top             =   120
         Visible         =   0   'False
         Width           =   420
      End
   End
   Begin GraphLib.Graph Graphe 
      Height          =   6915
      Left            =   4680
      TabIndex        =   0
      Top             =   420
      Width           =   11535
      _Version        =   65536
      _ExtentX        =   20346
      _ExtentY        =   12197
      _StockProps     =   96
      BorderStyle     =   1
      Background      =   0
      GraphStyle      =   4
      GraphTitle      =   "Production"
      GraphType       =   6
      LabelEvery      =   100
      Labels          =   0
      LegendStyle     =   1
      NumPoints       =   20
      NumSets         =   3
      PrintStyle      =   1
      RandomData      =   0
      YAxisMax        =   1
      YAxisMin        =   1
      ColorData       =   3
      ColorData[0]    =   12
      ColorData[1]    =   11
      ColorData[2]    =   10
      ExtraData       =   0
      ExtraData[]     =   0
      FontFamily      =   4
      FontSize        =   4
      FontSize[0]     =   60
      FontSize[1]     =   50
      FontSize[2]     =   50
      FontSize[3]     =   60
      FontStyle       =   4
      GraphData       =   3
      GraphData[]     =   20
      GraphData[0,0]  =   0
      GraphData[0,1]  =   0
      GraphData[0,2]  =   0
      GraphData[0,3]  =   0
      GraphData[0,4]  =   0
      GraphData[0,5]  =   0
      GraphData[0,6]  =   0
      GraphData[0,7]  =   0
      GraphData[0,8]  =   0
      GraphData[0,9]  =   0
      GraphData[0,10] =   0
      GraphData[0,11] =   0
      GraphData[0,12] =   0
      GraphData[0,13] =   0
      GraphData[0,14] =   0
      GraphData[0,15] =   0
      GraphData[0,16] =   0
      GraphData[0,17] =   0
      GraphData[0,18] =   0
      GraphData[0,19] =   0
      GraphData[1,0]  =   0
      GraphData[1,1]  =   0
      GraphData[1,2]  =   0
      GraphData[1,3]  =   0
      GraphData[1,4]  =   0
      GraphData[1,5]  =   0
      GraphData[1,6]  =   0
      GraphData[1,7]  =   0
      GraphData[1,8]  =   0
      GraphData[1,9]  =   0
      GraphData[1,10] =   0
      GraphData[1,11] =   0
      GraphData[1,12] =   0
      GraphData[1,13] =   0
      GraphData[1,14] =   0
      GraphData[1,15] =   0
      GraphData[1,16] =   0
      GraphData[1,17] =   0
      GraphData[1,18] =   0
      GraphData[1,19] =   0
      GraphData[2,0]  =   0
      GraphData[2,1]  =   0
      GraphData[2,2]  =   0
      GraphData[2,3]  =   0
      GraphData[2,4]  =   0
      GraphData[2,5]  =   0
      GraphData[2,6]  =   0
      GraphData[2,7]  =   0
      GraphData[2,8]  =   0
      GraphData[2,9]  =   0
      GraphData[2,10] =   0
      GraphData[2,11] =   0
      GraphData[2,12] =   0
      GraphData[2,13] =   0
      GraphData[2,14] =   0
      GraphData[2,15] =   0
      GraphData[2,16] =   0
      GraphData[2,17] =   0
      GraphData[2,18] =   0
      GraphData[2,19] =   0
      LabelText       =   0
      LegendText      =   0
      PatternData     =   3
      PatternData[0]  =   1
      PatternData[1]  =   1
      PatternData[2]  =   1
      SymbolData      =   0
      XPosData        =   0
      XPosData[]      =   0
   End
End
Attribute VB_Name = "FVisualisationGraphesProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre de visualisation des graphes de la production
' Nom                    : FVisualisationGraphesProduction.frm
' Date de création : 17/10/2011
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const TITRE_FENETRE As String = "VISUALISATION DES GRAPHES DE PRODUCTION"
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- énumérations privées ---
Private Enum TYPES_GRAPHES
    TG_TENSION_ET_INTENSITE_PHASES_1_A_3 = 0
    TG_TENSION_ET_INTENSITE_PHASE_4 = 1
    TG_TEMPERATURE_PHASES_1_A_3 = 2
    TG_TEMPERATURE_PHASE_4 = 3
End Enum

'--- types privées ---

'--- renseignements d'un graphe ---
Private Type RenseignementsGraphe
    CheminEtNomGraphe As String                   'chemin et nom complet du graphe
    NumFicheProduction As String                     'n° de la fiche de production
End Type

'--- variables privées ---
Private PremiereActivation As Boolean
Private InterdireEvenements As Boolean                                    'pour interdire certains évènements

'--- tableaux privés ---

'--- variables publiques ---
Public NumFenetre As Long                                                          'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBEchellesDilatees_Click()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- effacement complet du graphe ---
    Graphe.DataReset = gphAllData

    '--- réaffichage avec les nouvelles échelles ---
    If OBChoixGraphe(TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3).value = True Then
        InitialisationGraphe TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3
        DessineGraphe TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3
    ElseIf OBChoixGraphe(TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4).value = True Then
        InitialisationGraphe TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
        DessineGraphe TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
    End If

End Sub

Private Sub CBImprimerGraphe_Click()
    On Error Resume Next
    ImpressionGraphe
End Sub

Private Sub CBImprimerGraphe_GotFocus()
    
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

Private Sub CBImprimerGraphe_LostFocus()
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

Private Sub Form_Activate()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- renseigne la fenêtre principale ---
    RenseigneFPrincipale
    
    '--- placement du focus ---
    If PremiereActivation = False Then
        Me.Refresh
        PremiereActivation = True
    End If
    
    '--- tracer du graphe ---
    DessineGraphe TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Initialisation générale du graphes
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InitialisationGraphe(ByVal TypeGraphe As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    With Graphe
    
        '--- commandes communes à tous les graphes ---
        .DataReset = 9                                 'initialisation
        .RandomData = 0                             'désactive le mode aléatoire
        .IndexStyle = 1                                 'style de l'index du tableau de données
        
        '--- fixe les tailles des caractères pour tous les textes ---
        .FontUse = gphAllText
        .FontSize = 90
        .FontFamily = gphSwiss
        .FontStyle = gphBold
        
        '--- titre du graphe ---
        .FontUse = gphGraphTitle
        .FontFamily = gphSwiss
        .FontSize = 110
        .FontStyle = gphUnderlined
        
        '--- labels des X et Y ---
        '.FontUse = gphLabels
        '.FontFamily = gphSwiss
        '.FontSize = 50
        '.FontStyle = gphBold
        
        '--- autres titres ---
        '.FontUse = gphOtherTitles
        '.FontFamily = gphSwiss
        '.FontSize = 90
        '.FontStyle = gphBoldItalic
        
        '--- jeux de données ---
        Select Case TypeGraphe
            
            Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3, _
                     TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                '--- tension et intensité ---
                .NumSets = 2
                .AutoInc = 0                                     'pas d'auto incrémentation
                .Labels = 1                                      'labels des X et Y activés
                .GraphType = 6                                'définition du type de graphique (lignes)
                .GraphStyle = 0                                'définition du style de graphique
                
                .ThickLines = gphLinesOn              'admet la largeur de lignes
                .PatternedLines = 1                        'fixe la largeur de lignes
                
                '--- couleurs ---
                .Background = 0                               'couleur de fond
                .Foreground = 15                             'couleur des textes
        
                '--- jeu de valeurs 1 ---
                .ThisSet = 1
                .YAxisStyle = 0                                 'méthode de l'échelle de l'axe des Y
                .ThickLines = 1                                 'admet la largeur de lignes
                .PatternData = 1                               'épaisseur de lignes
                .ColorData = 12
                .LegendText = "U (V)"
                
                '--- jeu de valeurs 1 ---
                .ThisSet = 2
                .YAxisStyle = 0                                 'méthode de l'échelle de l'axe des Y
                .ThickLines = 1                                 'admet la largeur de lignes
                .PatternData = 1                               'épaisseur de lignes
                .ColorData = 11
                .LegendText = "I (A)"
        
                '--- titres abscisses / ordonnées ---
                If CBEchellesDilatees.value = vbUnchecked Then
                    .LeftTitle = "U et I" & vbLf & "1V=100A"
                Else
                    .LeftTitle = "U et I" & vbLf & "dilatées"
                End If
                .BottomTitle = "Temps (secondes)"
                
        
            Case TYPES_GRAPHES.TG_TEMPERATURE_PHASES_1_A_3, _
                     TYPES_GRAPHES.TG_TEMPERATURE_PHASE_4
                '--- température ---
                .NumSets = 1
                .AutoInc = 0                                      'pas d'auto incrémentation
                .Labels = 1                                       'labels des X et Y activés
                .GraphType = 6                                'définition du type de graphique (lignes)
                .GraphStyle = 0                                'définition du style de graphique
                
                '--- couleurs ---
                .Background = 0                               'couleur de fond
                .Foreground = 15                              'couleur des textes

                '--- jeu de valeurs ---
                .ThisSet = 1
                
                .YAxisStyle = 2                                 'méthode de l'échelle de l'axe des Y
                .YAxisTicks = 20
                .YAxisMin = 150
                .YAxisMax = 250
                
                .ThickLines = 1                                 'admet la largeur de lignes
                .PatternData = 1                                'épaisseur de lignes
                .ColorData = 10
                .LegendText = "t"
                
                '--- titres abscisses / ordonnées ---
                .LeftTitle = "Température" & vbLf & "°C"
                .BottomTitle = "Temps (secondes)"
                
            Case Else
                .NumSets = 1
        
        End Select
        
        .DrawMode = 3                    'mode de traçage du graphe (provoque le rafraichissement)
    
    End With
        
    '--- grisage du bouton impression ---
    With CBImprimerGraphe
        .Enabled = False
        .Refresh
    End With
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Imprime le graphe dessiné à l'écran
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ImpressionGraphe()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- changement de couleurs pour l'impression ---
    With Graphe
        .Background = 15               'couleur de fond
        .Foreground = 0                  'couleur des textes
        .DrawMode = 3                    'mode de traçage du graphe (provoque le rafraichissement)
        .Refresh
    End With

    '--- transfert des paramètres à l'imprimante ---
    Printer.Orientation = 2
    Printer.PaintPicture Graphe.Picture, 0, 0, 16000, 10000
    'Printer.CurrentY = 10000
    Printer.Print "TECAL - VERBRUGGE"
    Printer.NewPage
    Printer.EndDoc
    
    '--- changement de couleurs pour l'écran ---
    With Graphe
        .Background = 0                  'couleur de fond
        .Foreground = 15                 'couleur des textes
        .DrawMode = 3                    'mode de traçage du graphe (provoque le rafraichissement)
        .Refresh
    End With
        
    '--- transfert des paramètres à l'imprimante ---
    'Printer.Zoom = 400
    'Printer.PaintPicture Graphe.Picture, 0, 0, 16000, 5000
    'Printer.PaintPicture Graphe.Picture, 0, 5000, 16000, 5000
    'Printer.Print "Document des établissements G. VERBRUGGE"
    'Printer.NewPage
    'Printer.EndDoc
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Dessine un graphe en fonction du type de graphe
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DessineGraphe(ByVal TypeGraphe As TYPES_GRAPHES)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim PremierPassage As Boolean
    
    Dim NumFic As Integer
    
    Dim a As Long, _
            LongueurFichier As Long
    Dim XSecondes As Long
    Dim NbrPointsGraphe As Long
    Dim NbrPointsATracer As Long
    Dim PositionPoint As Long
    Dim PointDepartTracer As Long
    
    Dim IFiche As Single, UFiche As Single, TempFiche As Single
    
    Dim UMini As Single, UMaxi As Single
    Dim IMini As Single, IMaxi As Single
    Dim TempMini As Single, TempMaxi As Single
    Dim CoefMoins As Single, CoefPlus As Single
    
    Dim DatePremierPoint As Date
    
    Dim CheminEtNomFichier As String
    Dim TitreGraphe As String
    Dim TTraçabilite As Traçabilite
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- initialisation du graphe ---
    InitialisationGraphe (TypeGraphe)
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- affectation du chemin complet du fichier de traçabilité ---
    With TRenseignementsGraphe
        CheminEtNomFichier = RepGraphesProductionServeur & "F" & Right(String(8, "0") & .NumFicheProduction, 8) & _
                                                                                                    "D" & Format(.DateEntreeEnLigne, "ddmmyyyy") & _
                                                                                                    "H" & Format(.DateEntreeEnLigne, "hhnnss") & _
                                                                                                    "R" & CStr(.NumRedresseur) & _
                                                                                                    ".TRA"
    End With
    
    '--- sortie directe si le nom du fichier est vide ---
    If CheminEtNomFichier = "" Then Exit Sub
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    If FileExist(CheminEtNomFichier) = True Then
    
        '--- calcul de la longueur du fichier ---
        LongueurFichier = FileLen(CheminEtNomFichier)
    
        '--- calcul du nombre de points ---
        If LongueurFichier >= Len(TTraçabilite) Then
            NbrPointsGraphe = LongueurFichier / Len(TTraçabilite)
        End If
    
        '--- contrôle ---
        If NbrPointsGraphe > 0 Then
        
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- contrôle sur le nombre de points ---
            If NbrPointsGraphe > NBR_POINTS_MAXI_TRACABILITE Then
                NbrPointsGraphe = NBR_POINTS_MAXI_TRACABILITE
            End If
            
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- ouverture du fichier ---
            NumFic = FreeFile
            Open CheminEtNomFichier For Random Shared As #NumFic Len = Len(TTraçabilite)
            
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- affectations par défauts ---
            Select Case TypeGraphe
                
                Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3, TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                    '--- tension et intensité ---
                    UMini = 99999!
                    UMaxi = -99999!
                    IMini = 99999!
                    IMaxi = -99999!
    
                Case TYPES_GRAPHES.TG_TEMPERATURE_PHASES_1_A_3, TYPES_GRAPHES.TG_TEMPERATURE_PHASE_4
                    '--- température ---
                    TempMini = 99999!
                    TempMaxi = -99999!
                
                Case Else
            
            End Select
            
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- affectation ---
            NbrPointsATracer = 0
            PointDepartTracer = 1
            
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- recherche des mini. et maxi. ---
            For a = 1 To NbrPointsGraphe
                
                '--- lecture de la fiche ---
                Get #NumFic, a, TTraçabilite
                
                With TTraçabilite
                    
                    Select Case TypeGraphe
                        
                        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                        Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3
                            '--- tension et intensité des phases 1 à 3 ---
                            If .NumPhase < 4 Then
                            
                                '--- affectation de la tension et du courant ---
                                UFiche = CSng(.Tension) / 10
                                IFiche = CSng(.Intensite)
                                    
                                '--- affectation des valeurs mini et maxi pour l'échelle du graphe ---
                                UMini = IIf(UFiche < UMini, UFiche, UMini)
                                UMaxi = IIf(UFiche > UMaxi, UFiche, UMaxi)
                                IMini = IIf(IFiche < IMini, IFiche, IMini)
                                IMaxi = IIf(IFiche > IMaxi, IFiche, IMaxi)
                                    
                                '--- incrémentation du nombre de points à tracer ---
                                Inc NbrPointsATracer
                            
                            Else
                                
                                '--- sortie directe après traitement des 3 phases ---
                                Exit For
                            
                            End If
                            
                        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                        Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                            '--- tension et intensité de la phase 4 ---
                            If .NumPhase = 4 Then
                                
                                '--- affectation du point de départ pour la phase 4 ---
                                If PremierPassage = False Then
                                    PointDepartTracer = a
                                    PremierPassage = True
                                End If
                                  
                                '--- affectation de la tension et du courant ---
                                UFiche = CSng(.Tension) / 10
                                IFiche = CSng(.Intensite)

                                '--- affectation des valeurs mini et maxi pour l'échelle du graphe ---
                                UMini = IIf(UFiche < UMini, UFiche, UMini)
                                UMaxi = IIf(UFiche > UMaxi, UFiche, UMaxi)
                                IMini = IIf(IFiche < IMini, IFiche, IMini)
                                IMaxi = IIf(IFiche > IMaxi, IFiche, IMaxi)
                            
                                '--- incrémentation du nombre de points à tracer ---
                                Inc NbrPointsATracer
                            
                            End If
                        
                        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                        Case TYPES_GRAPHES.TG_TEMPERATURE_PHASES_1_A_3
                            '--- température des phases 1 à 3 ---
                            If .NumPhase < 4 Then
                                
                                '--- affectation de la température ---
                                TempFiche = .Temperature
                                
                                '--- affectation des valeurs mini et maxi pour l'échelle du graphe ---
                                TempMini = IIf(TempFiche < TempMini, TempFiche, TempMini)
                                TempMaxi = IIf(TempFiche > TempMaxi, TempFiche, TempMaxi)
                                
                                '--- incrémentation du nombre de points à tracer ---
                                Inc NbrPointsATracer
                            
                            Else
                                
                                '--- sortie directe après traitement des 3 phases ---
                                Exit For
                            
                            End If
                        
                        '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        
                        Case TYPES_GRAPHES.TG_TEMPERATURE_PHASE_4
                            '--- température de la phase 4 ---
                            If .NumPhase = 4 Then
                                
                                '--- affectation du point de départ pour la phase 4 ---
                                If PremierPassage = False Then
                                    PointDepartTracer = a
                                    PremierPassage = True
                                End If
                                
                                '--- affectation de la température ---
                                TempFiche = .Temperature
                                
                                '--- affectation des valeurs mini et maxi pour l'échelle du graphe ---
                                TempMini = IIf(TempFiche < TempMini, TempFiche, TempMini)
                                TempMaxi = IIf(TempFiche > TempMaxi, TempFiche, TempMaxi)
                                
                                '--- incrémentation du nombre de points à tracer ---
                                Inc NbrPointsATracer
                            
                            End If
                    
                        Case Else
                    
                    End Select
                    
                    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                    '--- affectation pour éviter les divisions par zéro ---
                    Select Case TypeGraphe
                        
                        Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3, TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                            If IMini = 0 Then IMini = 0.01
                            If IMaxi = 0 Then IMaxi = 0.01
                            If UMini = 0 Then UMini = 0.01
                            If UMaxi = 0 Then UMaxi = 0.01
                        
                        Case TYPES_GRAPHES.TG_TEMPERATURE_PHASES_1_A_3, TYPES_GRAPHES.TG_TEMPERATURE_PHASE_4
                            If TempMini = 0 Then TempMini = 0.01
                            If TempMaxi = 0 Then TempMaxi = 0.01
                        
                        Case Else
                    End Select
                
                    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                    '--- calcul des coefficients ---
                    Select Case TypeGraphe
                        
                        Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3, TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                            CoefPlus = Abs(IIf(IMaxi >= UMaxi, IMaxi / UMaxi, UMaxi / IMaxi))
                            CoefMoins = Abs(IIf(IMini <= UMini, IMini / UMini, UMini / IMini))
                        
                        Case Else
                    End Select
                    
                    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                End With
                
                '--- traitement des autres événements ---
                DoEvents
            
            Next a

            '***************************************************************************************************************************
            '                                                                           Tracer du graphe
            '***************************************************************************************************************************
            
            '--- lecture du premier point ---
            Get #NumFic, PointDepartTracer, TTraçabilite
            DatePremierPoint = TTraçabilite.DateDuPoint

            With Graphe
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- affichage des titres et textes ---
                With TRenseignementsGraphe
                    
                    Select Case TypeGraphe
                        
                        Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3, TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                            TitreGraphe = "U, I en anodisation"
                        
                        Case TYPES_GRAPHES.TG_TEMPERATURE_PHASES_1_A_3, TYPES_GRAPHES.TG_TEMPERATURE_PHASE_4
                            TitreGraphe = "Température en anodisation"
                        
                        Case Else
                    End Select
                    
                    '--- affectation du titre du graphe ---
                    TitreGraphe = TitreGraphe & " - " & .NumFicheProduction
                    
                    '--- affectation du titre dans le graphe ---
                    Graphe.GraphTitle = TitreGraphe
                
                End With
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- point d'origine ---
                .NumPoints = NbrPointsATracer
                PositionPoint = 0
                .XPosData = 0
                .GraphData = 0
                
                '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                
                '--- lecture du fichier ---
                For a = PointDepartTracer To NbrPointsGraphe

                    '--- lecture ---
                    Get #NumFic, a, TTraçabilite

                    '--- affectation ---
                    XSecondes = DateDiff("s", DatePremierPoint, TTraçabilite.DateDuPoint)
            
                    '--- tracé du graphe ---
                    If XSecondes >= 0 Then

                        '--- calcul des coefficients ---
                        Select Case TypeGraphe
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASES_1_A_3
                                '--- tension et intensité des phases 1 à 3 ---
                                If TTraçabilite.NumPhase < 4 Then
                        
                                    '--- point de travail ---
                                    Inc PositionPoint
                                    .ThisPoint = PositionPoint
                                    
                                    '--- tension ---
                                    .ThisSet = 1
                                    .XPosData = XSecondes
                                        
                                    '--- tracer de la tension sur le graphe ---
                                    If CBEchellesDilatees.value = vbUnchecked Then
                                        .GraphData = TTraçabilite.Tension * 10
                                    Else
                                        If IMaxi >= UMaxi Then
                                            .GraphData = TTraçabilite.Tension / 10 * CoefPlus
                                        Else
                                            .GraphData = TTraçabilite.Tension / 10
                                        End If
                                    End If
                                
                                    '--- intensité ---
                                    .ThisSet = 2
                                    .XPosData = XSecondes
                                  
                                    '--- tracer de l'intensité sur le graphe ---
                                    If CBEchellesDilatees.value = vbUnchecked Then
                                        .GraphData = TTraçabilite.Intensite
                                    Else
                                        If IMaxi >= UMaxi Then
                                            .GraphData = TTraçabilite.Intensite
                                        Else
                                            .GraphData = TTraçabilite.Intensite * CoefPlus
                                        End If
                                    End If
                                    
                                Else
                                
                                    '--- sortie directe après traitement des 3 phases ---
                                    Exit For
                                
                                End If
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case TYPES_GRAPHES.TG_TENSION_ET_INTENSITE_PHASE_4
                                '--- tension et intensité de la phase 4 ---
                                If TTraçabilite.NumPhase = 4 Then
                        
                                    '--- point de travail ---
                                    Inc PositionPoint
                                    .ThisPoint = PositionPoint
                                    
                                    '--- tension ---
                                    .ThisSet = 1
                                    .XPosData = XSecondes
                                        
                                    '--- tracer de la tension sur le graphe ---
                                    If CBEchellesDilatees.value = vbUnchecked Then
                                        .GraphData = TTraçabilite.Tension * 10
                                    Else
                                        If IMaxi >= UMaxi Then
                                            .GraphData = TTraçabilite.Tension / 10 * CoefPlus
                                        Else
                                            .GraphData = TTraçabilite.Tension / 10
                                        End If
                                    End If
                                
                                    '--- intensité ---
                                    .ThisSet = 2
                                    .XPosData = XSecondes
                                    
                                    '--- tracer de l'intensité sur le graphe ---
                                    If CBEchellesDilatees.value = vbUnchecked Then
                                        .GraphData = TTraçabilite.Intensite
                                    Else
                                        If IMaxi >= UMaxi Then
                                            .GraphData = TTraçabilite.Intensite
                                        Else
                                            .GraphData = TTraçabilite.Intensite * CoefPlus
                                        End If
                                    End If
                                
                                End If
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case TYPES_GRAPHES.TG_TEMPERATURE_PHASES_1_A_3
                                '--- température des phases 1 à 3 ---
                                If TTraçabilite.NumPhase < 4 Then
                        
                                    '--- point de travail ---
                                    Inc PositionPoint
                                    .ThisPoint = PositionPoint
                                    
                                    '--- température ---
                                    .ThisSet = 1
                                    .XPosData = XSecondes
                                    .GraphData = TTraçabilite.Temperature
                                
                                Else
                                
                                    '--- sortie directe après traitement des 3 phases ---
                                    Exit For
                                
                                End If
                            
                            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                            
                            Case TYPES_GRAPHES.TG_TEMPERATURE_PHASE_4
                                '--- température de la phase 4 ---
                                If TTraçabilite.NumPhase = 4 Then
                        
                                    '--- point de travail ---
                                    Inc PositionPoint
                                    .ThisPoint = PositionPoint
                                    
                                    '--- température ---
                                    .ThisSet = 1
                                    .XPosData = XSecondes
                                    .GraphData = TTraçabilite.Temperature
                                
                                End If
                            
                            Case Else
                        
                        End Select

                    End If
            
                    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                    '--- tracer complet du graphe ---
                    .DrawMode = 3                   'mode de traçage du graphe (provoque le rafraichissement)
            
                    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
                    '--- traitement des autres événements ---
                    DoEvents
                
                Next a
            
            End With
                    
            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- fermeture du fichier ---
            Close #NumFic

            '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            '--- permettre l'impression ---
            With CBImprimerGraphe
                .Enabled = True
                .Refresh
            End With
        
        End If
    
    Else
    
        '--- affichage du message d'erreur ---
        MessageErreur TITRE_MESSAGES, MESSAGE_123
    
    End If
    
End Sub

Private Sub Form_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- positionnement du choix des mesures ---
    FChoixMesures.Height = Me.ScaleHeight - PBRenseignementsFenetre.Height - PBBoutons.Height - 2 * Screen.TwipsPerPixelY
    
    '--- redimensionnement du graphe ---
    With Me
        Graphe.Left = FChoixMesures.Width + 7 * Screen.TwipsPerPixelX
        Graphe.Top = PBRenseignementsFenetre.Height + Screen.TwipsPerPixelY
        Graphe.Width = .ScaleWidth - Graphe.Left - 2 * Screen.TwipsPerPixelX
        Graphe.Height = .ScaleHeight - PBRenseignementsFenetre.Height - PBBoutons.Height - 5 * Screen.TwipsPerPixelY
    End With

End Sub

Private Sub OBChoixGraphe_Click(Index As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- traçage du graphe ---
    DessineGraphe Index

End Sub

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- déclaration ---
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBImprimerGraphe.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBImprimerGraphe.Width
    
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
' Entrées :
' Retours :
' Détails  :
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
    PBBoutons.Picture = ImgFondDesBoutons

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

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Décharge la fenêtre
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub DechargeFenetre()
    
    '--- aiguillage en cas d'erreur ---
    On Error Resume Next
    
    '--- affectation ---
    PremiereActivation = False
    
    '--- curseur souris par défaut ---
    SourisEnAttente False

    '--- neutralisation du timer ---
    
    '--- déchargement de la fenêtre ---
    Me.Visible = False
    Unload Me
    Set OccFVisualisationGraphesProduction = Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Change le curseur de la souris en fonction de l'attente
' Entrées : AttenteOuiNon -> TRUE   = Curseur en forme de sablier
'                                             FALSE = Curseur par défaut
' Retours :
' Détails  :
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




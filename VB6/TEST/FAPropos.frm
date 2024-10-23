VERSION 5.00
Begin VB.Form FAPropos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5340
   ClientLeft      =   2490
   ClientTop       =   3765
   ClientWidth     =   6690
   ClipControls    =   0   'False
   Icon            =   "FAPropos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PBBoutons 
      Align           =   2  'Align Bottom
      DrawStyle       =   2  'Dot
      DrawWidth       =   16891
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   6630
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4245
      Width           =   6690
      Begin VB.CommandButton CBInfosSysteme 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Infos ..."
         DownPicture     =   "FAPropos.frx":058A
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
         Left            =   3540
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAPropos.frx":0C8C
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "Infos &système..."
         ToolTipText     =   " Informations du système "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1395
      End
      Begin VB.CommandButton CBQuitter 
         BackColor       =   &H00FFFFFF&
         Cancel          =   -1  'True
         Caption         =   "&Quitter"
         DownPicture     =   "FAPropos.frx":138E
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
         Left            =   5100
         MaskColor       =   &H00FF00FF&
         Picture         =   "FAPropos.frx":1A90
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   " Quitter cette fenêtre "
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   1395
      End
      Begin VB.Shape SFocus 
         BorderColor     =   &H000000FF&
         BorderWidth     =   5
         Height          =   345
         Left            =   120
         Top             =   180
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.PictureBox PBLogo 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ClipControls    =   0   'False
      Height          =   2715
      Left            =   60
      Picture         =   "FAPropos.frx":2192
      ScaleHeight     =   2612.01
      ScaleMode       =   0  'User
      ScaleWidth      =   2768.521
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   2880
   End
   Begin VB.Label LLibelles 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AVERTISSEMENT"
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
      Height          =   255
      Left            =   60
      TabIndex        =   9
      Top             =   3060
      Width           =   6555
   End
   Begin VB.Label LLongueur 
      Caption         =   "Longueur"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3060
      TabIndex        =   5
      Tag             =   "Version"
      Top             =   480
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   60
      X2              =   6600
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Label LDescription 
      Caption         =   "Description de l'application"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   3060
      TabIndex        =   4
      Tag             =   "Description de l'application"
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Label LTitre 
      Caption         =   "Titre"
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
      Height          =   255
      Left            =   3060
      TabIndex        =   3
      Tag             =   "Titre de l'application"
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label LVersion 
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3060
      TabIndex        =   2
      Tag             =   "Version"
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label LAvertissement 
      Caption         =   "Avertissement : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   60
      TabIndex        =   1
      Tag             =   "Avertissement: ..."
      Top             =   3420
      Width           =   6495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   3
      X1              =   60
      X2              =   6600
      Y1              =   2940
      Y2              =   2940
   End
End
Attribute VB_Name = "FAPropos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : Fenêtre présentant les caractéristiques du logiciel
' Nom                    : FAPropos.frm
' Date de création : 29/03/1999
' Détails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---
Private Const KEY_ALL_ACCESS = &H2003F                 'reg Key - Options de sécurité ...
Private Const HKEY_LOCAL_MACHINE = &H80000002  'reg Key - Types de ROOT...
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1                                               'chaîne Unicode terminée par 0
Private Const REG_DWORD = 4                                      '32-bit number
Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO = "PATH"

Private Const TITRE_FENETRE As String = "A propos de ..."
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- déclarations privées ---
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'--- variables publiques ---
Public NumFenetre As Long    'numéro de la fenêtre lorsqu'elle devient active

Private Sub CBInfosSysteme_Click()
    InformationsSysteme
End Sub

Private Sub CBInfosSysteme_GotFocus()
    
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

Private Sub CBInfosSysteme_LostFocus()
    On Error Resume Next
    SFocus.Visible = False
End Sub

Private Sub CBQuitter_Click()
    Unload Me
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

End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim LongueurProgrammeOctets As Long
    Dim LongueurProgrammeMo As Single
    
    '--- divers sur la fenêtre ---
    Centrefenetre Me, TITRE_FENETRE
    PBBoutons.Picture = ImgFondDesBoutons
    
    '--- affectation ---
    LongueurProgrammeOctets = FileLen(App.Path & "\" & App.Title & ".exe")
    LongueurProgrammeMo = CSng(LongueurProgrammeOctets) / (1024! ^ 2)
    
    '--- affichage ---
    LTitre.Caption = App.Title & ".exe"
    LLongueur.Caption = Format(LongueurProgrammeOctets, "### ### ###") & " octets (" & _
                                      Format(LongueurProgrammeMo, "##0.0") & " Mo)"
    LVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    LDescription.Caption = "Gestion de production assistée par ordinateur (G.P.A.O)." & vbCrLf & vbCrLf & _
                                          "TECAL VERBRUGGE - Anodisation"
    LAvertissement.Caption = "Ce logiciel est protégé par les lois du copyright. Toute reproduction ou distribution, partielle ou totale, sans l'accord écrit de l'auteur, est strictement interdite."

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle     : Appel du programme des informations système
' Détails :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InformationsSysteme()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes privées ---
    Const MESSAGE_ERREURS As String = "Les informations sur le système ne sont pas disponibles pour l'instant"
    
    '--- déclaration ---
    Dim CheminProgSysInfo As String

    '--- lecture dans la base de registres du chemin\nom du programme des informations système ---
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, CheminProgSysInfo) Then
        
    '--- lecture dans la base de registres du chemin du programme des informations système ---
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, CheminProgSysInfo) Then
                
        '--- valider l'existence d'une version du fichier 32 bits connue ---
        If (Dir(CheminProgSysInfo & "\MSINFO32.EXE") <> "") Then
            CheminProgSysInfo = CheminProgSysInfo & "\MSINFO32.EXE"
        Else
            bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_ERREURS)
        End If
        
    Else
        
        '--- erreur - entrée de la base de registres introuvable ---
        bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_ERREURS)
            
    End If

    '--- appel du programme des informations système ---
    Call Shell(CheminProgSysInfo, vbNormalFocus)
    
    Exit Sub

'--- gestion des erreurs ---
GestionErreurs:
    bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_ERREURS)
        
End Sub

Public Function GetKeyValue(ByRef KeyRoot As Long, _
                                               ByRef KeyName As String, _
                                               ByRef SubKeyRef As String, _
                                               ByRef KeyVal As String) As Boolean
        
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
        
    '--- déclaration ---
    Dim a As Long                                           'compteur de boucle
    Dim CodeDeRetour As Long
    Dim hKey As Long                                     'pointeur vers une clé de registre ouvert
    Dim hDepth As Long
    Dim KeyValType As Long                          'type de données d'une clé de registre
    Dim tmpVal As String                                 'stockage temp. pour une valeur de clé de registre
    Dim KeyValSize As Long                           'taille de la variable clé de registre
    
    '--- ouvrir RegKey sous KeyRoot {HKEY_LOCAL_MACHINE...} / gestion des erreurs ------
    CodeDeRetour = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) 'ouvrir clé de registre
    If CodeDeRetour <> ERROR_SUCCESS Then GoTo GestionErreurs
        
    '--- affectation ---
    tmpVal = String$(1024, 0)       'allouer l'espace pour la variable
    KeyValSize = 1024                 'marquer la taille de la variable
        
    '--- extraire la valeur de clé de registre  / gestion des erreurs ------
    CodeDeRetour = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize) 'lire/créer validation de clé
    If CodeDeRetour <> ERROR_SUCCESS Then GoTo GestionErreurs

    '--- extraction de la chaine ---
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then       'Win95 termine les chaînes par 0...
        tmpVal = Left(tmpVal, KeyValSize - 1)               'Null atteint, extraire de la chaîne
    Else                                                                       'WinNT ne termine pas les chaînes par 0...
        tmpVal = Left(tmpVal, KeyValSize)                    '0 non trouvé, extraire chaîne uniquement
    End If
        
    '--- déterminer le type de la valeur de la clé pour la convertir ---
    Select Case KeyValType                                                'rechercher types de données...
        
        Case REG_SZ                                                             'type de données de clé de registre String
            KeyVal = tmpVal                                                      'copier valeur de la chaîne
        
        Case REG_DWORD                                                    'type de données de clé de registre Double Word
            For a = Len(tmpVal) To 1 Step -1                             'convertir chaque bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, a, 1)))   'construire valeur caractère par caractère
            Next a
            KeyVal = Format$("&h" + KeyVal)                            'convertir Double Word en String
    
    End Select
        
    '--- affectation ---
    GetKeyValue = True                                     'renvoyer réussite
    CodeDeRetour = RegCloseKey(hKey)          'fermer la clé de registre
        
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    KeyVal = ""                                                   'affecter la chaîne vide à la valeur de retour
    GetKeyValue = False                                   'renvoyer échec
    CodeDeRetour = RegCloseKey(hKey)         'fermer la clé de registre

End Function

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBInfosSysteme.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBInfosSysteme.Width
    
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

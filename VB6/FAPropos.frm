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
         Tag             =   "Infos &syst�me..."
         ToolTipText     =   " Informations du syst�me "
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
         ToolTipText     =   " Quitter cette fen�tre "
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
' R�le                    : Fen�tre pr�sentant les caract�ristiques du logiciel
' Nom                    : FAPropos.frm
' Date de cr�ation : 29/03/1999
' D�tails                :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- constantes priv�es ---
Private Const KEY_ALL_ACCESS = &H2003F                 'reg Key - Options de s�curit� ...
Private Const HKEY_LOCAL_MACHINE = &H80000002  'reg Key - Types de ROOT...
Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1                                               'cha�ne Unicode termin�e par 0
Private Const REG_DWORD = 4                                      '32-bit number
Private Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Private Const gREGVALSYSINFOLOC = "MSINFO"
Private Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Private Const gREGVALSYSINFO = "PATH"

Private Const TITRE_FENETRE As String = "A propos de ..."
Private Const TITRE_MESSAGES As String = TITRE_FENETRE

'--- d�clarations priv�es ---
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

'--- variables publiques ---
Public NumFenetre As Long    'num�ro de la fen�tre lorsqu'elle devient active

Private Sub CBInfosSysteme_Click()
    InformationsSysteme
End Sub

Private Sub CBInfosSysteme_GotFocus()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�placement du focus sur le bouton ---
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

    '--- d�placement du focus sur le bouton ---
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
    
    '--- renseigne la fen�tre principale ---
    RenseigneFPrincipale

End Sub

Private Sub Form_Load()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim LongueurProgrammeOctets As Long
    Dim LongueurProgrammeMo As Single
    
    '--- divers sur la fen�tre ---
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
    LDescription.Caption = "Gestion de production assist�e par ordinateur (G.P.A.O)." & vbCrLf & vbCrLf & _
                                          "TECAL VERBRUGGE - Anodisation"
    LAvertissement.Caption = "Ce logiciel est prot�g� par les lois du copyright. Toute reproduction ou distribution, partielle ou totale, sans l'accord �crit de l'auteur, est strictement interdite."

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le     : Appel du programme des informations syst�me
' D�tails :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub InformationsSysteme()
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    Const MESSAGE_ERREURS As String = "Les informations sur le syst�me ne sont pas disponibles pour l'instant"
    
    '--- d�claration ---
    Dim CheminProgSysInfo As String

    '--- lecture dans la base de registres du chemin\nom du programme des informations syst�me ---
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, CheminProgSysInfo) Then
        
    '--- lecture dans la base de registres du chemin du programme des informations syst�me ---
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, CheminProgSysInfo) Then
                
        '--- valider l'existence d'une version du fichier 32 bits connue ---
        If (Dir(CheminProgSysInfo & "\MSINFO32.EXE") <> "") Then
            CheminProgSysInfo = CheminProgSysInfo & "\MSINFO32.EXE"
        Else
            bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_ERREURS)
        End If
        
    Else
        
        '--- erreur - entr�e de la base de registres introuvable ---
        bidon = MessageErreur(TITRE_MESSAGES, MESSAGE_ERREURS)
            
    End If

    '--- appel du programme des informations syst�me ---
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
        
    '--- d�claration ---
    Dim a As Long                                           'compteur de boucle
    Dim CodeDeRetour As Long
    Dim hKey As Long                                     'pointeur vers une cl� de registre ouvert
    Dim hDepth As Long
    Dim KeyValType As Long                          'type de donn�es d'une cl� de registre
    Dim tmpVal As String                                 'stockage temp. pour une valeur de cl� de registre
    Dim KeyValSize As Long                           'taille de la variable cl� de registre
    
    '--- ouvrir RegKey sous KeyRoot {HKEY_LOCAL_MACHINE...} / gestion des erreurs ------
    CodeDeRetour = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) 'ouvrir cl� de registre
    If CodeDeRetour <> ERROR_SUCCESS Then GoTo GestionErreurs
        
    '--- affectation ---
    tmpVal = String$(1024, 0)       'allouer l'espace pour la variable
    KeyValSize = 1024                 'marquer la taille de la variable
        
    '--- extraire la valeur de cl� de registre  / gestion des erreurs ------
    CodeDeRetour = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize) 'lire/cr�er validation de cl�
    If CodeDeRetour <> ERROR_SUCCESS Then GoTo GestionErreurs

    '--- extraction de la chaine ---
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then       'Win95 termine les cha�nes par 0...
        tmpVal = Left(tmpVal, KeyValSize - 1)               'Null atteint, extraire de la cha�ne
    Else                                                                       'WinNT ne termine pas les cha�nes par 0...
        tmpVal = Left(tmpVal, KeyValSize)                    '0 non trouv�, extraire cha�ne uniquement
    End If
        
    '--- d�terminer le type de la valeur de la cl� pour la convertir ---
    Select Case KeyValType                                                'rechercher types de donn�es...
        
        Case REG_SZ                                                             'type de donn�es de cl� de registre String
            KeyVal = tmpVal                                                      'copier valeur de la cha�ne
        
        Case REG_DWORD                                                    'type de donn�es de cl� de registre Double Word
            For a = Len(tmpVal) To 1 Step -1                             'convertir chaque bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, a, 1)))   'construire valeur caract�re par caract�re
            Next a
            KeyVal = Format$("&h" + KeyVal)                            'convertir Double Word en String
    
    End Select
        
    '--- affectation ---
    GetKeyValue = True                                     'renvoyer r�ussite
    CodeDeRetour = RegCloseKey(hKey)          'fermer la cl� de registre
        
    Exit Function

'--- gestion des erreurs ---
GestionErreurs:
    KeyVal = ""                                                   'affecter la cha�ne vide � la valeur de retour
    GetKeyValue = False                                   'renvoyer �chec
    CodeDeRetour = RegCloseKey(hKey)         'fermer la cl� de registre

End Function

Private Sub PBBoutons_Resize()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- calculs de l'emplacement des boutons ---
    CBQuitter.Left = PBBoutons.ScaleWidth - MARGES.M_BORD_DROIT - CBQuitter.Width
    CBInfosSysteme.Left = CBQuitter.Left - MARGES.M_ENTRE_BOUTONS - CBInfosSysteme.Width
    
    '--- recalcul du focus apr�s d�placement ---
    With SFocus
        If .Visible = True Then
            .Left = ActiveControl.Left
            .Top = ActiveControl.Top
            .Height = ActiveControl.Height
            .Width = ActiveControl.Width
        End If
    End With

End Sub

Attribute VB_Name = "MFichiers"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DE GESTION DES FICHIERS SUR DISQUE
' Nom                    : MFichiers.bas
' Date de cr�ation : 09/03/2001
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
    (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
    ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    
    
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
    (ByVal lpApplicationName As String, _
     ByVal lpKeyName As String, _
     ByVal lpString As String, _
     ByVal lpFileName As String) As Long

Public Function EcrireDansIni(Section As String, Cle As String, Valeur As String)
    Dim Resultat As Long
    Resultat = WritePrivateProfileString(Section, Cle, Valeur, CONFIG_FILE)
    
    If Resultat = 0 Then
        MsgBox "Erreur lors de l'�criture dans le fichier .ini", vbCritical
    End If
End Function


Function StringToBoolean(ByVal inputString As String) As Boolean
    ' Convertit une cha�ne en bool�en selon certaines r�gles
    Select Case LCase(Trim(inputString)) ' Convertit en minuscule et supprime les espaces
        Case "true", "1", "yes", "oui"
            StringToBoolean = True
        Case "false", "0", "no", "non"
            StringToBoolean = False
        Case Else
            ' Si la cha�ne ne correspond � aucun cas, on retourne False par d�faut
            ' Vous pouvez lever une erreur ici si n�cessaire
            StringToBoolean = False
    End Select
End Function



'--- constantes priv�es ---

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauvegarde des graphes de la tra�abilit� des redresseurs
' D�tails  :
' Entr�es :               NumCharge -> Num�ro de charge
'                NumFicheProduction -> Num�ro de la fiche de production
'                  DateEntreeEnLigne -> Date entr�e en ligne de la charge
'                       NumRedresseur -> Num�ro du redresseur
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveTra�abiliteRedresseurs(ByVal NumCharge As Integer, _
                                                                           ByVal NumFicheProduction As String, _
                                                                           ByVal DateEntreeEnLigne As Date, _
                                                                           ByVal NumRedresseur As Integer) As String

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NomFichierTra�abilite As String                     'nom du fichier de tra�abilit�
    Dim NomFichierDefinitif As String                          'nom du fichier d�finitif
    
    '--- affectation du nom de fichier de tra�abilit� ---
    NomFichierTra�abilite = RepGraphesProductionLocal & "AnalyseRedresseurCharge" & Right("0" & NumCharge, 2) & ".FIC"
    
    
    
    '--- affectation du nom de fichier d�finitif ---
    NomFichierDefinitif = RepGraphesProductionServeur & "F" & Right(String(8, "0") & NumFicheProduction, 8) & _
                                                                                              "D" & Format(DateEntreeEnLigne, "ddmmyyyy") & _
                                                                                              "H" & Format(DateEntreeEnLigne, "hhnnss") & _
                                                                                              "R" & CStr(NumRedresseur) & _
                                                                                              ".TRA"
    
    
    '--- copie du fichier ---
    
    If FileExist(NomFichierTra�abilite) = False Then
        Call Log("Fichier: " & NomFichierTra�abilite & " introuvable !")
        MessageErreur "Erreur cr�ation graphe ", "Fichier: " & NomFichierTra�abilite & " introuvable !" & vbCrLf
    End If

    
    If FileExist(NomFichierDefinitif) = False Then
        If FolderExists(RepGraphesProductionServeur) = False Then
            MessageErreur "Erreur cr�ation graphe ", "Erreur de cr�ation du fichier: " & NomFichierDefinitif & vbCrLf & "Le dossier " & RepGraphesProductionServeur & " n'existe pas." & vbCrLf
        Else
            FileCopy NomFichierTra�abilite, NomFichierDefinitif
        End If
        
    End If
                
    '--- destruction du fichier de tra�abilit� ---
    If FileExist(NomFichierTra�abilite) = True Then
        'Kill NomFichierTra�abilite
        Open NomFichierTra�abilite For Output As 1
        Close 1
    End If
    
    Exit Function

GestionErreurs:
    SauveTra�abiliteRedresseurs = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement complet du programmateur cyclique
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeProgCyclique() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- d�claration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            NumFic As Integer
    Static CptTemps As Integer
    Dim DateFichier As Date
    Static DerniereDateFichier As Date
    Dim CheminComplet As String

    '--- affectation ---
    ChargeProgCyclique = ""
    CheminComplet = RepFicAnodisation & FIC_PROG_CYCLIQUE

    If FileExist(CheminComplet) = True Then

        '--- affectation ---
        DateFichier = FileDateTime(CheminComplet)

        '--- comptage du temps ---
        If DateFichier <> DerniereDateFichier And DerniereDateFichier <> Empty Then
            Inc CptTemps
        End If

        If DerniereDateFichier = Empty Or CptTemps >= TEMPS_VALIDITE_FICHIER Then

            '--- affichage du type de t�che ---
            AfficheTypeTache ("Chargement du programmateur cyclique")

            '--- affectation ---
            NumFic = FreeFile(1)

            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic

            '--- lecture des donn�es ---
            For a = 1 To NBR_JOURS_PROG_CYCLIQUE
                For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
                    With TProgCyclique(a, b)
                        Input #NumFic, .TypeDeJournee
                        Input #NumFic, Bidon
                        For c = 1 To NBR_TOPS_POSSIBLES
                            Input #NumFic, .TTopsDebutPompe(c), .TTopsFinPompe(c), .TCyclesPompe(c)
                        Next c
                        Input #NumFic, Bidon
                        For c = 1 To NBR_TOPS_POSSIBLES
                            Input #NumFic, .TTopsDebutChauffage(c), .TTopsFinChauffage(c), .TModesChauffage(c)
                        Next c
                    End With
                Next b
            Next a

            '--- fermeture du fichier ---
            Close #NumFic

            '--- affectation ---
            CptTemps = 0
            DerniereDateFichier = DateFichier

            '--- affichage du type de t�che ---
            AfficheTypeTache ("")

        End If

    Else
    
        '--- fichier introuvable ---
        ChargeProgCyclique = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If

    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeProgCyclique = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement complet des param�tres de la ligne
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeParametresLigne() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- d�claration ---
    Dim a As Integer, _
            NumFic As Integer
    Static CptTemps As Integer
    Dim DateFichier As Date
    Static DerniereDateFichier As Date
    Dim CheminComplet As String

    '--- affectation ---
    ChargeParametresLigne = ""
    CheminComplet = RepFicAnodisation & FIC_PARAMETRES_LIGNE

    If FileExist(CheminComplet) = True Then

        '--- affectation ---
        DateFichier = FileDateTime(CheminComplet)

        '--- comptage du temps ---
        If DateFichier <> DerniereDateFichier And DerniereDateFichier <> Empty Then
            Inc CptTemps
        End If

        If DerniereDateFichier = Empty Or CptTemps >= TEMPS_VALIDITE_FICHIER Then

            '--- affichage du type de t�che ---
            AfficheTypeTache ("Chargement des param�tres de la ligne")

            '--- affectation ---
            NumFic = FreeFile(1)

            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic

            '--- lecture des donn�es ---
            For a = REDRESSEURS.R_C13 To REDRESSEURS.R_C16
                With TEtatsRedresseurs(a)
                    Input #NumFic, Bidon
                    'Input #NumFic, .TNumDefauts.DefautGeneral
                End With
            Next a

            '--- fermeture du fichier ---
            Close #NumFic

            '--- affectation ---
            CptTemps = 0
            DerniereDateFichier = DateFichier

            '--- affichage du type de t�che ---
            AfficheTypeTache ("")

        End If

    Else
    
        '--- fichier introuvable ---
        ChargeParametresLigne = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If

    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeParametresLigne = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauvegarde compl�te du programmateur cyclique
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveProgCyclique() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- d�claration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            NumFic As Integer
    Dim LibelleJournee As String

    '--- affectation ---
    SauveProgCyclique = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("Sauvegarde du programmateur cyclique")

    '--- affectation ---
    NumFic = FreeFile(1)

    '--- ouverture et �criture du fichier ---
   Open RepFicAnodisation & FIC_PROG_CYCLIQUE For Output Shared As #NumFic

    '--- enregistrement ---
    For a = 1 To NBR_JOURS_PROG_CYCLIQUE
        For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
            With TProgCyclique(a, b)

                '--- affectation sur le libell� de la journ�e ---
                Select Case .TypeDeJournee
                    Case JOURNEES_TYPES.J_ARRET: LibelleJournee = " - Arr�t"
                    Case JOURNEES_TYPES.J_TRAVAIL: LibelleJournee = " - Travail"
                    Case JOURNEES_TYPES.J_VEILLE: LibelleJournee = " - Veille"
                    Case JOURNEES_TYPES.J_REPRISE: LibelleJournee = " - Reprise"
                    Case Else: LibelleJournee = ""
                End Select

                '--- �criture sur le disque ---
                Write #NumFic, .TypeDeJournee
                Write #NumFic, "Journ�e " & CStr(a) & LibelleJournee & ", cuve " & TEtatsCuves(b).DefinitionCuve.NomCuve & ", pompe"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutPompe(c), .TTopsFinPompe(c), .TCyclesPompe(c)
                Next c
                Write #NumFic, "Journ�e " & CStr(a) & LibelleJournee & ", cuve " & TEtatsCuves(b).DefinitionCuve.NomCuve & ", chauffage"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutChauffage(c), .TTopsFinChauffage(c), .TModesChauffage(c)
                Next c

            End With
        Next b
    Next a

    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")

    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveProgCyclique = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement des annexes (ventilation, etc ...)
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeAnnexes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv�es ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"

    '--- d�claration ---
    Dim a As Integer, _
            NumFic As Integer
    Static CptTemps As Integer
    Dim DateFichier As Date
    Static DerniereDateFichier As Date
    Dim CheminComplet As String
    
    '--- affectation ---
    ChargeAnnexes = ""
    CheminComplet = RepFicAnodisation & FIC_ANNEXES
        
    If FileExist(CheminComplet) = True Then
    
        '--- affectation ---
        DateFichier = FileDateTime(CheminComplet)
    
        '--- comptage du temps ---
        If DateFichier <> DerniereDateFichier And DerniereDateFichier <> Empty Then
            Inc CptTemps
        End If
        
        If DerniereDateFichier = Empty Or CptTemps >= TEMPS_VALIDITE_FICHIER Then
        
            '--- affichage du type de t�che ---
            AfficheTypeTache ("Chargement des annexes")
        
            '--- affectation ---
            NumFic = FreeFile(1)
    
            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic
    
            '--- lecture des donn�es ---
            With TEtatsAnnexes
                Input #NumFic, Bidon
                Input #NumFic, .ModeEVBrillantage
                Input #NumFic, Bidon
                Input #NumFic, .PeriodiciteEVBrillantage
                Input #NumFic, Bidon
                Input #NumFic, .TempsMarcheEVBrillantage
            End With
    
            '--- fermeture du fichier ---
            Close #NumFic
  
            '--- affectation ---
            CptTemps = 0
            DerniereDateFichier = DateFichier
  
            '--- affichage du type de t�che ---
            AfficheTypeTache ("")
        
        End If
    
    Else
    
        '--- fichier introuvable ---
        ChargeAnnexes = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeAnnexes = CStr(Err.Number)
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauvegarde compl�te des annexes (ventilation, etc ...)
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveAnnexes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim a As Integer, _
            NumFic As Integer
            
    '--- affectation ---
    SauveAnnexes = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("Sauvegarde des annexes")
        
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et �criture du fichier ---
   Open RepFicAnodisation & FIC_ANNEXES For Output Shared As #NumFic
    
    '--- enregistrement ---
    With TEtatsAnnexes
        Write #NumFic, "Mode de l'�lectro-vanne d'air dans le bain de brillantage"
        Write #NumFic, .ModeEVBrillantage
        Write #NumFic, "P�riodicit� de mise en marche de l'�lectro-vanne d'air dans le bain de brillantage"
        Write #NumFic, .PeriodiciteEVBrillantage
        Write #NumFic, "Temps de marche de l'�lectro-vanne d'air dans le bain de brillantage"
        Write #NumFic, .TempsMarcheEVBrillantage
    End With
    
    '--- fermeture du fichier ---
    Close #NumFic
  
    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveAnnexes = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement complet de la r�gulation
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeRegulation() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv�es ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"

    '--- d�claration ---
    Dim a As Integer, _
            NumFic As Integer
    Static CptTemps As Integer
    Dim DateFichier As Date
    Static DerniereDateFichier As Date
    Dim CheminComplet As String
    
    '--- affectation ---
    ChargeRegulation = ""
    CheminComplet = RepFicAnodisation & FIC_REGULATION
    
    If FileExist(CheminComplet) = True Then
    
        '--- affectation ---
        DateFichier = FileDateTime(CheminComplet)
        
        '--- comptage du temps ---
        If DateFichier <> DerniereDateFichier And DerniereDateFichier <> Empty Then
            Inc CptTemps
        End If
        
        If DerniereDateFichier = Empty Or CptTemps >= TEMPS_VALIDITE_FICHIER Then
    
            '--- affichage du type de t�che ---
            AfficheTypeTache ("Chargement de la r�gulation")
 
            '--- affectation ---
            NumFic = FreeFile(1)
    
            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic
                
            '--- lecture des donn�es ---
            For a = LBound(TEtatsCuves()) To UBound(TEtatsCuves())
                With TEtatsCuves(a).Temperatures
                    Input #NumFic, Bidon
                    Input #NumFic, .TempVeille, .TempProduction
                    Input #NumFic, .EcartInferieurRegul, .EcartSuperieurRegul, .EcartInferieurAlarme, .EcartSuperieurAlarme
                End With
            Next a
    
            '--- fermeture du fichier ---
            Close #NumFic
  
            '--- affectation ---
            CptTemps = 0
            DerniereDateFichier = DateFichier
    
            '--- affichage du type de t�che ---
            AfficheTypeTache ("")
        
        End If
    
    Else
    
        '--- fichier introuvable ---
        ChargeRegulation = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If

    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeRegulation = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauvegarde compl�te de la r�gulation
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveRegulation() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim a As Integer, _
            NumFic As Integer
            
    '--- affectation ---
    SauveRegulation = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("Sauvegarde de la r�gulation")
    
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et �criture du fichier ---
   Open RepFicAnodisation & FIC_REGULATION For Output Shared As #NumFic
            
    '--- enregistrement ---
    For a = LBound(TEtatsCuves()) To UBound(TEtatsCuves())
        With TEtatsCuves(a).Temperatures
            Write #NumFic, "R�gulation cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve
            Write #NumFic, .TempVeille, .TempProduction,
            Write #NumFic, .EcartInferieurRegul, .EcartSuperieurRegul, .EcartInferieurAlarme, .EcartSuperieurAlarme
        End With
    Next a
    
    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveRegulation = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement complet des journ�es types
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeJourneesTypes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv�es ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- d�claration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            NumFic As Integer
    Static CptTemps As Integer
    Dim DateFichier As Date
    Static DerniereDateFichier As Date
    Dim CheminComplet As String
    
    '--- affectation ---
    ChargeJourneesTypes = ""
    CheminComplet = RepFicAnodisation & FIC_JOURNEES_TYPES
    
    If FileExist(CheminComplet) = True Then
    
        '--- affectation ---
        DateFichier = FileDateTime(CheminComplet)
        
        '--- comptage du temps ---
        If DateFichier <> DerniereDateFichier And DerniereDateFichier <> Empty Then
            Inc CptTemps
        End If
        
        If DerniereDateFichier = Empty Or CptTemps >= TEMPS_VALIDITE_FICHIER Then
    
            '--- affichage du type de t�che ---
            AfficheTypeTache ("Chargement des journ�es types")
            
            '--- affectation ---
            NumFic = FreeFile(1)
    
            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic
    
            '--- lecture des donn�es ---
            For a = LBound(TJourneesTypes()) To UBound(TJourneesTypes())
                For b = JOURNEES_TYPES.J_ARRET To JOURNEES_TYPES.J_REPRISE
                    With TJourneesTypes(a, b)
                        Input #NumFic, Bidon
                        For c = 1 To NBR_TOPS_POSSIBLES
                            Input #NumFic, .TTopsDebutPompe(c), .TTopsFinPompe(c), .TCyclesPompe(c)
                        Next c
                        Input #NumFic, Bidon
                        For c = 1 To NBR_TOPS_POSSIBLES
                            Input #NumFic, .TTopsDebutChauffage(c), .TTopsFinChauffage(c), .TModesChauffage(c)
                        Next c
                    End With
                Next b
            Next a
    
            '--- fermeture du fichier ---
            Close #NumFic

            '--- affectation ---
            CptTemps = 0
            DerniereDateFichier = DateFichier

            '--- affichage du type de t�che ---
            AfficheTypeTache ("")
        
        End If
    
    Else
    
        '--- fichier introuvable ---
        ChargeJourneesTypes = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If

    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeJourneesTypes = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauvegarde compl�te des journ�es types
' D�tails  :
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveJourneesTypes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            NumFic As Integer
    Dim LibelleJournee As String
             
    '--- affectation ---
    SauveJourneesTypes = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache ("Sauvegarde des journ�es types")
    
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et �criture du fichier ---
   Open RepFicAnodisation & FIC_JOURNEES_TYPES For Output Shared As #NumFic
    
    '--- enregistrement ---
    For a = LBound(TJourneesTypes()) To UBound(TJourneesTypes())
        For b = JOURNEES_TYPES.J_ARRET To JOURNEES_TYPES.J_REPRISE
            With TJourneesTypes(a, b)
                                
                '--- affectation sur la journ�e d'arr�t ---
                If b = 0 Then
                    .TTopsDebutPompe(1) = "XXXXXXXX000000"
                    .TTopsFinPompe(1) = "XXXXXXXX235959"
                    .TCyclesPompe(1) = 0
                    .TTopsDebutChauffage(1) = "XXXXXXXX000000"
                    .TTopsFinChauffage(1) = "XXXXXXXX235959"
                    .TModesChauffage(1) = 0
                    For c = 2 To NBR_TOPS_POSSIBLES
                        .TTopsDebutPompe(c) = ""
                        .TTopsFinPompe(c) = ""
                        .TCyclesPompe(c) = 0
                        .TTopsDebutChauffage(c) = ""
                        .TTopsFinChauffage(c) = ""
                        .TModesChauffage(c) = 0
                    Next c
                End If
                
                '--- affectation sur le libell� de la journ�e ---
                Select Case b
                    Case JOURNEES_TYPES.J_ARRET: LibelleJournee = "d'arr�t"
                    Case JOURNEES_TYPES.J_TRAVAIL: LibelleJournee = "de travail"
                    Case JOURNEES_TYPES.J_VEILLE: LibelleJournee = "de veille"
                    Case JOURNEES_TYPES.J_REPRISE: LibelleJournee = "de reprise"
                    Case Else: LibelleJournee = ""
                End Select
                
                '--- enregistrement ---
                Write #NumFic, "Cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve & ", journ�e " & LibelleJournee & ", pompe"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutPompe(c), .TTopsFinPompe(c), .TCyclesPompe(c)
                Next c
                Write #NumFic, "Cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve & ", journ�e " & LibelleJournee & ", chauffage"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutChauffage(c), .TTopsFinChauffage(c), .TModesChauffage(c)
                Next c
            
            End With
        Next b
    Next a
    
    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de t�che ---
    AfficheTypeTache ("")
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveJourneesTypes = CStr(Err.Number)
    
End Function




Public Function GetConnectionString(categorie As String, champ As String) As String
    Dim buffer As String * 255
    Dim filePath As String
    filePath = CONFIG_FILE
    GetPrivateProfileString categorie, champ, "", buffer, 255, filePath
    GetConnectionString = Left$(buffer, InStr(buffer, Chr$(0)) - 1)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Chargement de la configuration
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeConfiguration() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante priv�es ---
    Const CONFIGURATION As String = "Configuration"
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- d�claration ---
    Dim NumFic As Integer
    Dim CheminComplet  As String
    Dim BDD As String
    
    '--- affectation ---
    ChargeConfiguration = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Chargement de la configuration"


    '--- lecture des valeurs dans la base de registres ---
    MemMenuPrincipalNavigateur = GetSetting(App.Title, CONFIGURATION, "M�moire du menu principal du navigateur", 0)
    MemSousMenuNavigateur = GetSetting(App.Title, CONFIGURATION, "M�moire du sous menu du navigateur", 0)
    RepLocalBD = GetSetting(App.Title, CONFIGURATION, "R�pertoire de la base de donn�es en local", "V:\")
    RepDistantBD = GetSetting(App.Title, CONFIGURATION, "R�pertoire de la base de donn�es en distant", "C:\Gv")
    ModeDeConnexion = GetSetting(App.Title, CONFIGURATION, "Mode de connexion", 0)
    MotDePasseDirection = GetSetting(App.Title, CONFIGURATION, "Mot de passe direction", "")
    MotDePasseDirection = DecodeMotDePasse(MotDePasseDirection)
    MotDePassePersonnel = GetSetting(App.Title, CONFIGURATION, "Mot de passe personnel", "")
    MotDePassePersonnel = DecodeMotDePasse(MotDePassePersonnel)

    
    PARAMETRES_CONNEXION_BD_CLIPPER_HF = GetConnectionString("database", "CLIPPER")
    
    'MODE_DECONNECTE = StringToBoolean(LCase(Trim(GetConnectionString("parametres", "MODE_DECONNECTE"))))
   
    If Environ("ANODISATION_TEST") = 1 Then
        'VM XP
        PARAMETRES_CONNEXION_BD_ANODISATION_SQL = GetConnectionString("database", "SQLEXPRESS")
        PROGRAMME_AVEC_AUTOMATE = False
    Else
        PROGRAMME_AVEC_AUTOMATE = True
        'BDD = GetConnectionString("parametres", "BDD")
        'Select Case BDD
        '    Case "PROD"
        '        PARAMETRES_CONNEXION_BD_ANODISATION_SQL = GetConnectionString("database", "SQLEXPRESS_PROD")
        '    Case "LOCAL"
        '        MsgBox ("BASE DE DONNEES LOCALE !")
        '        PARAMETRES_CONNEXION_BD_ANODISATION_SQL = GetConnectionString("database", "SQLEXPRESS_LOCAL")
        '    Case Else
        '         MsgBox ("Erreur lecture config.ini pour la BDD")
        'End Select
     
    End If
    
  
    'MsgBox ("PARAMETRES_CONNEXION_BD_ANODISATION_SQL=" & PARAMETRES_CONNEXION_BD_ANODISATION_SQL)
    'MsgBox ("Environ(ANODISATION_TEST) " & Environ("ANODISATION_TEST"))
  
    Dim b As Boolean
    
    b = StringToBoolean(LCase(Trim(GetConnectionString("GestionRegulation", "AppareillageConcerne"))))
         
    With VManipsGestionRegulation
            .AppareillageConcerne = b
            .CyclesPompe = Val(GetConnectionString("GestionRegulation", "CyclesPompe"))
            .ModesChauffage = Val(GetConnectionString("GestionRegulation", "ModesChauffage"))
    End With
    
    b = StringToBoolean(LCase(Trim(GetConnectionString("ManipsProgCyclique", "AppareillageConcerne"))))
    With VManipsProgCyclique
        .AppareillageConcerne = b
        .CyclesPompe = Val(GetConnectionString("ManipsProgCyclique", "CyclesPompe"))
        .ModesChauffage = Val(GetConnectionString("ManipsProgCyclique", "ModesChauffage"))
    End With
   
    
    SHOW_LOGS = StringToBoolean(LCase(Trim(GetConnectionString("parametres", "LOGS"))))
    
    MemDateProgCyclique = GetConnectionString("parametres", "MemDateProgCyclique")
    
    Exit Function

GestionErreurs:
   
    ChargeConfiguration = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauvegarde de la configuration
' Entr�es :
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveConfiguration() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes priv�es ---
    Const CONFIGURATION As String = "Configuration"
    
    '--- d�claration ---
    Dim NumFic As Integer
    Dim CheminComplet As String
    
    '--- affectation ---
    SauveConfiguration = ""
    
    '--- affichage du type de t�che ---
    AfficheTypeTache "Sauvegarde de la configuration"
                    
    '--- enregistrement des valeurs dans la base de registres ---
    SaveSetting App.Title, CONFIGURATION, "R�pertoire de la base de donn�es en local", RepLocalBD
    SaveSetting App.Title, CONFIGURATION, "R�pertoire de la base de donn�es en distant", RepDistantBD
    SaveSetting App.Title, CONFIGURATION, "Mode de connexion", ModeDeConnexion
    SaveSetting App.Title, CONFIGURATION, "Mot de passe direction", CodeMotDePasse(MotDePasseDirection)
    SaveSetting App.Title, CONFIGURATION, "Mot de passe personnel", CodeMotDePasse(MotDePassePersonnel)
    SaveSetting App.Title, CONFIGURATION, "Suppression des mots de passe", SuppressionMotsDePasse
    SaveSetting App.Title, CONFIGURATION, "Indice de la prestation par d�faut", IndicePrestationParDefaut
    SaveSetting App.Title, CONFIGURATION, "Libell� de la prestation par d�faut", LibellePrestationParDefaut
    SaveSetting App.Title, CONFIGURATION, "Nombre de lignes maxi. � extraire", NbrLignesMaxiAExtraire
    SaveSetting App.Title, CONFIGURATION, "Temps de compensation d'anodisation", TempsCompensationAnodisationMinutes
       
    '--- LE FICHIER DE CONFIGURATION DOIT SE TROUVER DANS LE REPERTOIRE DU PROGRAMME ---
    
    
    EcrireDansIni "parametres", "MemDateProgCyclique", MemDateProgCyclique
   
  

    '--- manipulations dans la fen�tre gestion de la r�gulation ---
    Write #NumFic, "Manipulations dans la fen�tre gestion de la r�gulation"
    With VManipsGestionRegulation
        EcrireDansIni "GestionRegulation", "AppareillageConcerne", CStr(.AppareillageConcerne)
        EcrireDansIni "GestionRegulation", "CyclesPompe", CStr(.CyclesPompe)
        EcrireDansIni "GestionRegulation", "ModesChauffage", CStr(.ModesChauffage)
    End With

    '--- manipulations dans la fen�tre du programmateur cyclique ---
    Write #NumFic, "Manipulations dans la fen�tre du programmateur cyclique"
    With VManipsProgCyclique
        EcrireDansIni "ManipsProgCyclique", "AppareillageConcerne", CStr(.AppareillageConcerne)
        EcrireDansIni "ManipsProgCyclique", "CyclesPompe", CStr(.CyclesPompe)
        EcrireDansIni "ManipsProgCyclique", "ModesChauffage", CStr(.ModesChauffage)
    End With

    

    '--- affichage du type de t�che ---
    AfficheTypeTache ""
    
    Exit Function

GestionErreurs:
    
    SauveConfiguration = CStr(Err.Number)
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Sauve l'�tats des postes
' D�tails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveEtatsPostes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- d�claration ---
    Dim a As Integer, _
            NumFic As Integer
    
    '--- affectation ---
    SauveEtatsPostes = ""
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC. Then Exit Function

    '--- affectation ---
    NumFic = FreeFile(1)

    '--- ouverture du fichier ---
    Open RepFicAnodisation & FIC_ETATS_POSTES For Output Shared As #NumFic
   
    '--- enregistrement ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a)
            
            '--- d�finition des postes ---
            Write #NumFic, .DefinitionPoste.NumPoste, .DefinitionPoste.NomPoste, .DefinitionPoste.LibellePoste
            Write #NumFic, .DefinitionPoste.AvecEgouttage, .DefinitionPoste.PresenceCouvercles, .DefinitionPoste.PresenceRedresseur, .DefinitionPoste.PresenceAgitationBain
            Write #NumFic, .DefinitionPoste.XAxePosteLigne, .DefinitionPoste.XAxePosteSynoptique
            Write #NumFic, .DefinitionPoste.XInferieurPosteSynoptique, .DefinitionPoste.YInferieurPosteSynoptique, .DefinitionPoste.XSuperieurPosteSynoptique, .DefinitionPoste.YSuperieurPosteSynoptique
            Write #NumFic, .DefinitionPoste.XInferieurLibellePosteSynoptique, .DefinitionPoste.YInferieurLibellePosteSynoptique, .DefinitionPoste.XSuperieurLibellePosteSynoptique, .DefinitionPoste.YSuperieurLibellePosteSynoptique

            '--- �tats du reste de la fiche ---
            Write #NumFic, .NumCharge, .Condamnation, .EtatsChariots
            Write #NumFic, .Alarmes
    
        End With
    Next a
    
    '--- fermeture du fichier ---
    Close #NumFic
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveEtatsPostes = CStr(Err.Number)

End Function



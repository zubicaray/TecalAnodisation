Attribute VB_Name = "MFichiers"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DE GESTION DES FICHIERS SUR DISQUE
' Nom                    : MFichiers.bas
' Date de création : 09/03/2001
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes privées ---

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Sauvegarde des graphes de la traçabilité des redresseurs
' Détails  :
' Entrées :               NumCharge -> Numéro de charge
'                NumFicheProduction -> Numéro de la fiche de production
'                  DateEntreeEnLigne -> Date entrée en ligne de la charge
'                       NumRedresseur -> Numéro du redresseur
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveTraçabiliteRedresseurs(ByVal NumCharge As Integer, _
                                                                           ByVal NumFicheProduction As String, _
                                                                           ByVal DateEntreeEnLigne As Date, _
                                                                           ByVal NumRedresseur As Integer) As String

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NomFichierTraçabilite As String                     'nom du fichier de traçabilité
    Dim NomFichierDefinitif As String                          'nom du fichier définitif
    
    '--- affectation du nom de fichier de traçabilité ---
    NomFichierTraçabilite = RepGraphesProductionLocal & "AnalyseRedresseurCharge" & Right("0" & NumCharge, 2) & ".FIC"
    
    
    
    '--- affectation du nom de fichier définitif ---
    NomFichierDefinitif = RepGraphesProductionServeur & "F" & Right(String(8, "0") & NumFicheProduction, 8) & _
                                                                                              "D" & Format(DateEntreeEnLigne, "ddmmyyyy") & _
                                                                                              "H" & Format(DateEntreeEnLigne, "hhnnss") & _
                                                                                              "R" & CStr(NumRedresseur) & _
                                                                                              ".TRA"
    
    
   

    
    '--- copie du fichier ---
    
    If FileExist(NomFichierTraçabilite) = False Then
        Call Log("Fichier: " & NomFichierTraçabilite & " introuvable !")
        MessageErreur "Erreur création graphe ", "Fichier: " & NomFichierTraçabilite & " introuvable !" & vbCrLf
    End If

    
    If FileExist(NomFichierDefinitif) = False Then
        If FolderExists(RepGraphesProductionServeur) = False Then
            MessageErreur "Erreur création graphe ", "Erreur de création du fichier: " & NomFichierDefinitif & vbCrLf & "Le dossier " & RepGraphesProductionServeur & " n'existe pas." & vbCrLf
        Else
            FileCopy NomFichierTraçabilite, NomFichierDefinitif
        End If
        
    End If
                
    '--- destruction du fichier de traçabilité ---
    If FileExist(NomFichierTraçabilite) = True Then
        'Kill NomFichierTraçabilite
        Open NomFichierTraçabilite For Output As 1
        Close 1
    End If
    
    Exit Function

GestionErreurs:
    SauveTraçabiliteRedresseurs = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Chargement complet du programmateur cyclique
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeProgCyclique() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes privées ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- déclaration ---
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

            '--- affichage du type de tâche ---
            AfficheTypeTache ("Chargement du programmateur cyclique")

            '--- affectation ---
            NumFic = FreeFile(1)

            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic

            '--- lecture des données ---
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

            '--- affichage du type de tâche ---
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
' Rôle      : Chargement complet des paramètres de la ligne
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeParametresLigne() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes privées ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- déclaration ---
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

            '--- affichage du type de tâche ---
            AfficheTypeTache ("Chargement des paramètres de la ligne")

            '--- affectation ---
            NumFic = FreeFile(1)

            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic

            '--- lecture des données ---
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

            '--- affichage du type de tâche ---
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
' Rôle      : Sauvegarde complète du programmateur cyclique
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveProgCyclique() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            NumFic As Integer
    Dim LibelleJournee As String

    '--- affectation ---
    SauveProgCyclique = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache ("Sauvegarde du programmateur cyclique")

    '--- affectation ---
    NumFic = FreeFile(1)

    '--- ouverture et écriture du fichier ---
   Open RepFicAnodisation & FIC_PROG_CYCLIQUE For Output Shared As #NumFic

    '--- enregistrement ---
    For a = 1 To NBR_JOURS_PROG_CYCLIQUE
        For b = CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION
            With TProgCyclique(a, b)

                '--- affectation sur le libellé de la journée ---
                Select Case .TypeDeJournee
                    Case JOURNEES_TYPES.J_ARRET: LibelleJournee = " - Arrêt"
                    Case JOURNEES_TYPES.J_TRAVAIL: LibelleJournee = " - Travail"
                    Case JOURNEES_TYPES.J_VEILLE: LibelleJournee = " - Veille"
                    Case JOURNEES_TYPES.J_REPRISE: LibelleJournee = " - Reprise"
                    Case Else: LibelleJournee = ""
                End Select

                '--- écriture sur le disque ---
                Write #NumFic, .TypeDeJournee
                Write #NumFic, "Journée " & CStr(a) & LibelleJournee & ", cuve " & TEtatsCuves(b).DefinitionCuve.NomCuve & ", pompe"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutPompe(c), .TTopsFinPompe(c), .TCyclesPompe(c)
                Next c
                Write #NumFic, "Journée " & CStr(a) & LibelleJournee & ", cuve " & TEtatsCuves(b).DefinitionCuve.NomCuve & ", chauffage"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutChauffage(c), .TTopsFinChauffage(c), .TModesChauffage(c)
                Next c

            End With
        Next b
    Next a

    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de tâche ---
    AfficheTypeTache ("")

    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveProgCyclique = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Chargement des annexes (ventilation, etc ...)
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeAnnexes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"

    '--- déclaration ---
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
        
            '--- affichage du type de tâche ---
            AfficheTypeTache ("Chargement des annexes")
        
            '--- affectation ---
            NumFic = FreeFile(1)
    
            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic
    
            '--- lecture des données ---
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
  
            '--- affichage du type de tâche ---
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
' Rôle      : Sauvegarde complète des annexes (ventilation, etc ...)
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveAnnexes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim a As Integer, _
            NumFic As Integer
            
    '--- affectation ---
    SauveAnnexes = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache ("Sauvegarde des annexes")
        
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et écriture du fichier ---
   Open RepFicAnodisation & FIC_ANNEXES For Output Shared As #NumFic
    
    '--- enregistrement ---
    With TEtatsAnnexes
        Write #NumFic, "Mode de l'électro-vanne d'air dans le bain de brillantage"
        Write #NumFic, .ModeEVBrillantage
        Write #NumFic, "Périodicité de mise en marche de l'électro-vanne d'air dans le bain de brillantage"
        Write #NumFic, .PeriodiciteEVBrillantage
        Write #NumFic, "Temps de marche de l'électro-vanne d'air dans le bain de brillantage"
        Write #NumFic, .TempsMarcheEVBrillantage
    End With
    
    '--- fermeture du fichier ---
    Close #NumFic
  
    '--- affichage du type de tâche ---
    AfficheTypeTache ("")
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveAnnexes = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Chargement complet de la régulation
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeRegulation() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"

    '--- déclaration ---
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
    
            '--- affichage du type de tâche ---
            AfficheTypeTache ("Chargement de la régulation")
 
            '--- affectation ---
            NumFic = FreeFile(1)
    
            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic
                
            '--- lecture des données ---
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
    
            '--- affichage du type de tâche ---
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
' Rôle      : Sauvegarde complète de la régulation
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveRegulation() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim a As Integer, _
            NumFic As Integer
            
    '--- affectation ---
    SauveRegulation = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache ("Sauvegarde de la régulation")
    
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et écriture du fichier ---
   Open RepFicAnodisation & FIC_REGULATION For Output Shared As #NumFic
            
    '--- enregistrement ---
    For a = LBound(TEtatsCuves()) To UBound(TEtatsCuves())
        With TEtatsCuves(a).Temperatures
            Write #NumFic, "Régulation cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve
            Write #NumFic, .TempVeille, .TempProduction,
            Write #NumFic, .EcartInferieurRegul, .EcartSuperieurRegul, .EcartInferieurAlarme, .EcartSuperieurAlarme
        End With
    Next a
    
    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de tâche ---
    AfficheTypeTache ("")
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveRegulation = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Chargement complet des journées types
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeJourneesTypes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- déclaration ---
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
    
            '--- affichage du type de tâche ---
            AfficheTypeTache ("Chargement des journées types")
            
            '--- affectation ---
            NumFic = FreeFile(1)
    
            '--- ouverture et lecture du fichier ---
            Open CheminComplet For Input Shared As #NumFic
    
            '--- lecture des données ---
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

            '--- affichage du type de tâche ---
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
' Rôle      : Sauvegarde complète des journées types
' Détails  :
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveJourneesTypes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim a As Integer, _
            b As Integer, _
            c As Integer, _
            NumFic As Integer
    Dim LibelleJournee As String
             
    '--- affectation ---
    SauveJourneesTypes = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache ("Sauvegarde des journées types")
    
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et écriture du fichier ---
   Open RepFicAnodisation & FIC_JOURNEES_TYPES For Output Shared As #NumFic
    
    '--- enregistrement ---
    For a = LBound(TJourneesTypes()) To UBound(TJourneesTypes())
        For b = JOURNEES_TYPES.J_ARRET To JOURNEES_TYPES.J_REPRISE
            With TJourneesTypes(a, b)
                                
                '--- affectation sur la journée d'arrêt ---
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
                
                '--- affectation sur le libellé de la journée ---
                Select Case b
                    Case JOURNEES_TYPES.J_ARRET: LibelleJournee = "d'arrêt"
                    Case JOURNEES_TYPES.J_TRAVAIL: LibelleJournee = "de travail"
                    Case JOURNEES_TYPES.J_VEILLE: LibelleJournee = "de veille"
                    Case JOURNEES_TYPES.J_REPRISE: LibelleJournee = "de reprise"
                    Case Else: LibelleJournee = ""
                End Select
                
                '--- enregistrement ---
                Write #NumFic, "Cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve & ", journée " & LibelleJournee & ", pompe"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutPompe(c), .TTopsFinPompe(c), .TCyclesPompe(c)
                Next c
                Write #NumFic, "Cuve " & TEtatsCuves(a).DefinitionCuve.NomCuve & ", journée " & LibelleJournee & ", chauffage"
                For c = 1 To NBR_TOPS_POSSIBLES
                    Write #NumFic, .TTopsDebutChauffage(c), .TTopsFinChauffage(c), .TModesChauffage(c)
                Next c
            
            End With
        Next b
    Next a
    
    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de tâche ---
    AfficheTypeTache ("")
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveJourneesTypes = CStr(Err.Number)
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Chargement du chemin de la base de données CLIPPER
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeCheminBDCLIPPER() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante privées ---
    Const CONFIGURATION As String = "Configuration"
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- déclaration ---
    Dim NumFic As Integer
    Dim CheminComplet  As String
    
    '--- affectation ---
    ChargeCheminBDCLIPPER = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache "Chargement du chemin de la base de données CLIPPER"

    '--- LE FICHIER DE CONFIGURATION DOIT SE TROUVER DANS LE REPERTOIRE DU PROGRAMME ---
    
    '--- affectation ---
    CheminComplet = App.Path & "\" & FIC_CONFIGURATION
    
    If FileExist(CheminComplet) = True Then
  
        '--- affectation ---
        NumFic = FreeFile(1)
    
        '--- ouverture et lecture du fichier ---
        Open CheminComplet For Input Shared As #NumFic

        '--- type de PC ---
        Input #NumFic, Bidon
        Input #NumFic, Bidon
        
        '--- programmateur cyclique ---
        Input #NumFic, Bidon
        Input #NumFic, Bidon
    
        '--- manipulations dans la fenêtre gestion de la régulation ---
        Input #NumFic, Bidon
        With VManipsGestionRegulation
            Input #NumFic, Bidon
            Input #NumFic, Bidon
            Input #NumFic, Bidon
        End With
    
        '--- manipulations dans la fenêtre du programmateur cyclique ---
        Input #NumFic, Bidon
        With VManipsProgCyclique
            Input #NumFic, Bidon
            Input #NumFic, Bidon
            Input #NumFic, Bidon
        End With
    
        '--- chemin des bains pour CLIPPER ---
        Input #NumFic, Bidon
        Input #NumFic, RepFicClipper
        

        '--- fermeture du fichier ---
        Close #NumFic
    
        '--- affichage du type de tâche ---
        AfficheTypeTache ("")
    
    Else
    
        '--- fichier introuvable ---
        ChargeCheminBDCLIPPER = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If
  
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeCheminBDCLIPPER = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Chargement de la configuration
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChargeConfiguration() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    '--- constante privées ---
    Const CONFIGURATION As String = "Configuration"
    Const CODE_ERREUR_FICHIER_INTROUVABLE As String = "53"
    
    '--- déclaration ---
    Dim NumFic As Integer
    Dim CheminComplet  As String
    
    '--- affectation ---
    ChargeConfiguration = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache "Chargement de la configuration"


    '--- lecture des valeurs dans la base de registres ---
    MemMenuPrincipalNavigateur = GetSetting(App.Title, CONFIGURATION, "Mémoire du menu principal du navigateur", 0)
    MemSousMenuNavigateur = GetSetting(App.Title, CONFIGURATION, "Mémoire du sous menu du navigateur", 0)
    RepLocalBD = GetSetting(App.Title, CONFIGURATION, "Répertoire de la base de données en local", "V:\")
    RepDistantBD = GetSetting(App.Title, CONFIGURATION, "Répertoire de la base de données en distant", "C:\Gv")
    ModeDeConnexion = GetSetting(App.Title, CONFIGURATION, "Mode de connexion", 0)
    MotDePasseDirection = GetSetting(App.Title, CONFIGURATION, "Mot de passe direction", "")
    MotDePasseDirection = DecodeMotDePasse(MotDePasseDirection)
    MotDePassePersonnel = GetSetting(App.Title, CONFIGURATION, "Mot de passe personnel", "")
    MotDePassePersonnel = DecodeMotDePasse(MotDePassePersonnel)
    'SuppressionMotsDePasse = GetSetting(App.Title, CONFIGURATION, "Suppression des mots de passe", True)
    'UniteMonetaire = GetSetting(App.Title, CONFIGURATION, "Unité monètaire (0=Francs français, 1=Euro)", 0)
    'IndicePrestationParDefaut = GetSetting(App.Title, CONFIGURATION, "Indice de la prestation par défaut", 0)
    'LibellePrestationParDefaut = GetSetting(App.Title, CONFIGURATION, "Libellé de la prestation par défaut", "CHROMAGE")
    'NbrLignesMaxiAExtraire = GetSetting(App.Title, CONFIGURATION, "Nombre de lignes maxi. à extraire", 0)
    'TempsCompensationAnodisationMinutes = GetSetting(App.Title, CONFIGURATION, "Temps de compensation d'anodisation", 0)
    
    '--- LE FICHIER DE CONFIGURATION DOIT SE TROUVER DANS LE REPERTOIRE DU PROGRAMME ---
    
    '--- affectation ---
    CheminComplet = App.Path & "\" & FIC_CONFIGURATION
    TypePC = TYPES_PC.PC_SUR_LIGNE
    If FileExist(CheminComplet) = True Then
  
        '--- affectation ---
        NumFic = FreeFile(1)
    
        '--- ouverture et lecture du fichier ---
        Open CheminComplet For Input Shared As #NumFic

        'MODE SECOURS
        Input #NumFic, Bidon
        Input #NumFic, varConfig
       
        If varConfig = 1 Then
            MODE_SECOURS = True
        Else
            If varConfig = 0 Then
                MODE_SECOURS = False
            Else
                MODE_SECOURS = False
                MsgBox ("Vous devez mettre 1 pour vrai, 0 pour faux")
            End If
        
        
        
            
        
        End If
       

        'CNX BDD ANODISATION
        Input #NumFic, Bidon
        Input #NumFic, varConfig
        
        If varConfig = TYPE_BDD_ANO.PROD Then
            PARAMETRES_CONNEXION_BD_ANODISATION_SQL = CST_PARAMETRES_CONNEXION_BD_ANODISATION_SQL
        End If
        If varConfig = TYPE_BDD_ANO.TEST Then
            PARAMETRES_CONNEXION_BD_ANODISATION_SQL = CST_PARAMETRES_CONNEXION_BD_ANODISATION_TEST_SQL
        End If
        
        
        'MsgBox (PARAMETRES_CONNEXION_BD_ANODISATION_SQL)
        'CNX BDD CLIPPER
        Input #NumFic, Bidon
        Input #NumFic, varConfig
        
        If varConfig = TYPE_BDD_ANO.PROD Then
            PARAMETRES_CONNEXION_BD_CLIPPER_HF = CST_PARAMETRES_CONNEXION_BD_CLIPPER_HF
        End If
        
        
        If varConfig = TYPE_BDD_CLIPPER.ACCESS_TEST Then
            PARAMETRES_CONNEXION_BD_CLIPPER_HF = CST_PARAMETRES_CONNEXION_BD_CLIPPER_TEST_ACCESS
        End If
        If varConfig = TYPE_BDD_ANO.TEST Then
            PARAMETRES_CONNEXION_BD_CLIPPER_HF = CST_PARAMETRES_CONNEXION_BD_CLIPPER_TEST_HF
        End If
        

        
        
        '--- programmateur cyclique ---
        Input #NumFic, Bidon
        Input #NumFic, MemDateProgCyclique
    
        '--- manipulations dans la fenêtre gestion de la régulation ---
        Input #NumFic, Bidon
        With VManipsGestionRegulation
            Input #NumFic, .AppareillageConcerne
            Input #NumFic, .CyclesPompe
            Input #NumFic, .ModesChauffage
        End With
    
        '--- manipulations dans la fenêtre du programmateur cyclique ---
        Input #NumFic, Bidon
        With VManipsProgCyclique
            Input #NumFic, .AppareillageConcerne
            Input #NumFic, .CyclesPompe
            Input #NumFic, .ModesChauffage
        End With
    
    
        
        '--- chemin des bains pour CLIPPER ---
        Input #NumFic, Bidon
        Input #NumFic, RepFicClipper
    
      
          '--- affichier les logs
        Input #NumFic, Bidon
        Input #NumFic, varConfig
      
        If varConfig = 1 Then
            ShowLog = True
        Else
            ShowLog = False
        End If
                
        Close #NumFic
    
    
     
    
        '--- affichage du type de tâche ---
        AfficheTypeTache ("")
    
    Else
    
        '--- fichier introuvable ---
        ChargeConfiguration = CODE_ERREUR_FICHIER_INTROUVABLE
    
    End If
  
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    ChargeConfiguration = CStr(Err.Number)

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Sauvegarde de la configuration
' Entrées :
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveConfiguration_OLD() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- constantes privées ---
    Const CONFIGURATION As String = "Configuration"
    
    '--- déclaration ---
    Dim NumFic As Integer
    Dim CheminComplet As String
    
    '--- affectation ---
    SauveConfiguration_OLD = ""
    
    '--- affichage du type de tâche ---
    AfficheTypeTache "Sauvegarde de la configuration"
                    
    '--- enregistrement des valeurs dans la base de registres ---
    SaveSetting App.Title, CONFIGURATION, "Répertoire de la base de données en local", RepLocalBD
    SaveSetting App.Title, CONFIGURATION, "Répertoire de la base de données en distant", RepDistantBD
    SaveSetting App.Title, CONFIGURATION, "Mode de connexion", ModeDeConnexion
    SaveSetting App.Title, CONFIGURATION, "Mot de passe direction", CodeMotDePasse(MotDePasseDirection)
    SaveSetting App.Title, CONFIGURATION, "Mot de passe personnel", CodeMotDePasse(MotDePassePersonnel)
    SaveSetting App.Title, CONFIGURATION, "Suppression des mots de passe", SuppressionMotsDePasse
    SaveSetting App.Title, CONFIGURATION, "Indice de la prestation par défaut", IndicePrestationParDefaut
    SaveSetting App.Title, CONFIGURATION, "Libellé de la prestation par défaut", LibellePrestationParDefaut
    SaveSetting App.Title, CONFIGURATION, "Nombre de lignes maxi. à extraire", NbrLignesMaxiAExtraire
    SaveSetting App.Title, CONFIGURATION, "Temps de compensation d'anodisation", TempsCompensationAnodisationMinutes
       
    '--- LE FICHIER DE CONFIGURATION DOIT SE TROUVER DANS LE REPERTOIRE DU PROGRAMME ---
    
    '--- affectation ---
    CheminComplet = App.Path & "\" & FIC_CONFIGURATION
    
    '--- affectation ---
    NumFic = FreeFile(1)
    
    '--- ouverture et écriture du fichier ---
    Open CheminComplet For Output Shared As #NumFic

    '--- type de PC ---
    Write #NumFic, "Indique le type de PC (1 = PC de la ligne d'anodisation, 2 = PC Entreprise, 3 = PC Distant)"
    Write #NumFic, TypePC
    
    '--- programmateur cyclique ---
    Write #NumFic, "Mémoire de la date pour changer le programmateur cyclique"
    Write #NumFic, MemDateProgCyclique

    '--- manipulations dans la fenêtre gestion de la régulation ---
    Write #NumFic, "Manipulations dans la fenêtre gestion de la régulation"
    With VManipsGestionRegulation
        Write #NumFic, .AppareillageConcerne
        Write #NumFic, .CyclesPompe
        Write #NumFic, .ModesChauffage
    End With

    '--- manipulations dans la fenêtre du programmateur cyclique ---
    Write #NumFic, "Manipulations dans la fenêtre du programmateur cyclique"
    With VManipsProgCyclique
        Write #NumFic, .AppareillageConcerne
        Write #NumFic, .CyclesPompe
        Write #NumFic, .ModesChauffage
    End With

    '--- chemin des bains pour CLIPPER ---
    Write #NumFic, "Chemin des bains pour CLIPPER"
    Write #NumFic, RepFicClipper
    
    '--- fermeture du fichier ---
    Close #NumFic

    '--- affichage du type de tâche ---
    AfficheTypeTache ""
    
    Exit Function

GestionErreurs:
    If NumFic > 0 Then Close #NumFic
    SauveConfiguration_OLD = CStr(Err.Number)
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Sauve l'états des postes
' Détails  :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SauveEtatsPostes() As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
    
    '--- déclaration ---
    Dim a As Integer, _
            NumFic As Integer
    
    '--- affectation ---
    SauveEtatsPostes = ""
    
    '--- analyse en fonction du PC ---
    'If TypePC <> TYPES_PC.PC_SUR_LIGNE Then Exit Function

    '--- affectation ---
    NumFic = FreeFile(1)

    '--- ouverture du fichier ---
    Open RepFicAnodisation & FIC_ETATS_POSTES For Output Shared As #NumFic
   
    '--- enregistrement ---
    For a = LBound(TEtatsPostes()) To UBound(TEtatsPostes())
        With TEtatsPostes(a)
            
            '--- définition des postes ---
            Write #NumFic, .DefinitionPoste.NumPoste, .DefinitionPoste.NomPoste, .DefinitionPoste.LibellePoste
            Write #NumFic, .DefinitionPoste.AvecEgouttage, .DefinitionPoste.PresenceCouvercles, .DefinitionPoste.PresenceRedresseur, .DefinitionPoste.PresenceAgitationBain
            Write #NumFic, .DefinitionPoste.XAxePosteLigne, .DefinitionPoste.XAxePosteSynoptique
            Write #NumFic, .DefinitionPoste.XInferieurPosteSynoptique, .DefinitionPoste.YInferieurPosteSynoptique, .DefinitionPoste.XSuperieurPosteSynoptique, .DefinitionPoste.YSuperieurPosteSynoptique
            Write #NumFic, .DefinitionPoste.XInferieurLibellePosteSynoptique, .DefinitionPoste.YInferieurLibellePosteSynoptique, .DefinitionPoste.XSuperieurLibellePosteSynoptique, .DefinitionPoste.YSuperieurLibellePosteSynoptique

            '--- états du reste de la fiche ---
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



VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInsertionClipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' //
' // CCopyProgress - class for copying with notification
' //

Option Explicit



' SZP 2023
Public Function insertionClipperPointageBackground(ByVal NumCharge As Integer)
    
    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs
  
    Dim lHeures As Double

    '--- déclaration ---
    Dim a As Integer, _
    NumAction As Integer, _
    LastPOCLEUNIK  As Long
    Dim Requete As String
    Dim ConnexionBDClipper As New ADODB.Connection
   
    Dim logClipper As Boolean
    logClipper = True
    
    
    
    'screen.MousePointer = vbHourglass
    Call Log("DEBUT INSERTION CLIPPER", logClipper)

    '--- ouverture de la connexion à la base de données Clipper 09 ---
    With ConnexionBDClipper
        .ConnectionString = PARAMETRES_CONNEXION_BD_CLIPPER_HF
        .CursorLocation = adUseClient
        .Mode = adModeReadWrite
        .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
        .Open
    End With
    Dim EnregistrementPoint As New ADODB.Recordset
    Dim EnregistrementId As New ADODB.Recordset
    Dim EnregistrementComplem1 As New ADODB.Recordset
    Dim EnregistrementComplem2 As New ADODB.Recordset
    Dim EnregistrementUPD As New ADODB.Recordset
  
    
    '--- fermeture de l'enregistrement / connexion ---
    
    
    
    'POCLEUNIK,
    ' & LastPOCLEUNIK
    With TEtatsCharges(NumCharge)
        
        'TMP
        lHeures = DateDiff("s", .DateEntreeEnLigne, .DateArriveeAuDechargement) / 3600
        Dim l2Heures As String
        l2Heures = Replace(lHeures, ",", ".")
        Dim dateIn As Long
        Dim dateOut As Long
        Dim heureIn As String
        Dim heureOut As String
        dateIn = Format(.DateEntreeEnLigne, "YYYYMMDD")
        dateOut = Format(.DateArriveeAuDechargement, "YYYYMMDD")
        
        heureIn = Format(CStr(.DateEntreeEnLigne), "hhnn")
        heureOut = Format(CStr(.DateArriveeAuDechargement), "hhnn")
        
        'LastPOCLEUNIK = LastPOCLEUNIK + 1
        Dim lastId As Long
        
        'Call Log("------------------------------------------------------------------------")
        Call Log("CLIPPER numéro de barre : " & .NumBarreInc & ", temps en heure:" & l2Heures, logClipper)
         
        Call Log("DateEntreeEnLigne : " & CStr(.DateEntreeEnLigne) & ", DateArriveeAuDechargement:" & CStr(.DateArriveeAuDechargement), logClipper)
        Call Log("heureIn : " & heureIn & ",heureOut:" & heureOut, logClipper)
        
        
               
        
        For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
            
            If (.TDetailsCharges(a).NumCommandeInterne > 0) Then
            
            
                Set EnregistrementPoint = Nothing
                Set EnregistrementId = Nothing
                Set EnregistrementComplem1 = Nothing
                Set EnregistrementComplem2 = Nothing
                Set EnregistrementUPD = Nothing
                
                Call Log("CLIPPER numéro de charge de la barre: " & a & ", GACLEUNIK:" & .TDetailsCharges(a).NumCommandeInterne, logClipper)
            
                Requete = "INSERT INTO POINT " & _
                "(GACLEUNIK,CODEEMP, NAF, FLAG,NBPIE,TPSPASSE,HEUREDEB,HEUREFIN,DAT,DATEFIN,COFRAIS) " & _
                "VALUES ( " & .TDetailsCharges(a).NumCommandeInterne & "," _
                & "'000'," & .TDetailsCharges(a).Naf & _
               ",'2'," & .TDetailsCharges(a).NbrPieces & "," & _
                l2Heures & ",'" & heureIn & "','" & heureOut & "'," & _
                dateIn & "," & dateOut & "," & _
               "'LIGTR'" & ")"
               
                
               
               
                With EnregistrementPoint
                    '--- lancement de la requête ---
                    .CursorLocation = adUseServer
                    .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic
                End With
                'EnregistrementPoint.Close
                
                Call Log("CLIPPER FIN INSERT INTO POINT ", logClipper)
                
                Requete = "SELECT MAX(IDPOINT)  AS IdAuto FROM Point  where GACLEUNIK=" & .TDetailsCharges(a).NumCommandeInterne
                With EnregistrementId
                    '--- lancement de la requête ---
                    .CursorLocation = adUseServer
                    .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic
                End With
                
                 Call Log("CLIPPER  FIN SELECT MAX(IDPOINT) ", logClipper)
                
                lastId = EnregistrementId.Fields(0)
                'Call Log("lastId: " & lastId)
                'EnregistrementId.Close
                
                
                'TMP
                Requete = "INSERT INTO COMPLEMS " & _
                "(CLEUNIK,COPAR, COMPLEMENT) " & _
                "VALUES ( " & lastId & "," _
                & "'POCPL01'," & .NumBarreInc _
                & ")"
                With EnregistrementComplem1
                    '--- lancement de la requête ---
                    .CursorLocation = adUseServer
                    .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic
                End With
                'EnregistrementComplem1.Close
                Call Log("CLIPPER FIN INSERT INTO COMPLEMS POCPL01 ", logClipper)
                               
                Requete = "INSERT INTO COMPLEMS " & _
                "(CLEUNIK,COPAR, COMPLEMENT) " & _
                "VALUES ( " & lastId & "," _
                & "'POCPL02'," & .TGammesAnodisation.NumGamme _
                & ")"
                With EnregistrementComplem2
                    '--- lancement de la requête ---
                    .CursorLocation = adUseServer
                    .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic
                End With
                'EnregistrementComplem2.Close
                
                Call Log("CLIPPER FIN INSERT INTO COMPLEMS POCPL02 ", logClipper)
                
                Call Log(".TDetailsCharges(a).TypeReparation= " & .TDetailsCharges(a).TypeReparation, logClipper)
               
                If (.TDetailsCharges(a).TypeReparation <> "") Then
                    Requete = "INSERT INTO COMPLEMS " & _
                    "(CLEUNIK,COPAR, COMPLEMENT) " & _
                    "VALUES ( " & lastId & "," _
                    & "'POCPL03','" & .TDetailsCharges(a).TypeReparation & "')"
                    With EnregistrementComplem2
                        '--- lancement de la requête ---
                        .CursorLocation = adUseServer
                        .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic
                    End With
                    'EnregistrementComplem2.Close
                    
                    Call Log("CLIPPER FIN INSERT INTO COMPLEMS POCPL03 ", logClipper)
                
                End If
                
              
                
                Requete = "UPDATE GAMME SET GA_HRES=GA_HRES-" & l2Heures & _
                ",GA_NBHR=GA_NBHR+" & l2Heures & _
                " WHERE GACLEUNIK= " & .TDetailsCharges(a).NumCommandeInterne
                With EnregistrementUPD
                    '--- lancement de la requête ---
                    .CursorLocation = adUseServer
                    .Open Requete, ConnexionBDClipper, adOpenStatic, adLockOptimistic
                End With
                'EnregistrementUPD.Close
                
                Call Log("CLIPPER FIN UPDATE GAMME ", logClipper)
                
            End If
            
            
        Next a
    
    
         
        
    End With
    

   
    
    '--- fermeture de l'enregistrement / connexion ---
   
     'Enregistrement.Close
     ConnexionBDClipper.Close
   
    
    'Set EnregistrementPoint = Nothing
    'Set EnregistrementId = Nothing
    'Set EnregistrementComplem1 = Nothing
    'Set EnregistrementComplem2 = Nothing
    'Set EnregistrementUPD = Nothing
    'Set ConnexionBDClipper = Nothing
    
    insertionClipperPointageBackground = True
    Call Log("FIN ET SSUCCES INSERTION CLIPPER")
    'Screen.MousePointer = vbNormal
   
    Exit Function

GestionErreurs:
    
    '--- valeur de retour ---
     
    Set EnregistrementPoint = Nothing
    Set EnregistrementId = Nothing
    Set EnregistrementComplem1 = Nothing
    Set EnregistrementComplem2 = Nothing
    Set EnregistrementUPD = Nothing
    
    MsgBox (CStr(Err.Description))
    Call Log("CLIPPER" & CStr(Err.Description))
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next

    ConnexionBDClipper.Close
    Set ConnexionBDClipper = Nothing
    insertionClipperPointageBackground = False

End Function
 
Public Sub insertionClipper(ByVal NumCharge As Integer)

    
    'Call Log("insertionClipperPointageBackground  " & NumCharge)

    insertionClipperPointageBackground NumCharge
  
End Sub






'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Enregistrement complet de la production d'une charge
' Entrées :                        NumCharge -> Numéro de la charge concernée
' Retours : EnregistrementProduction -> "" = pas d'incident sinon numéro de l'erreur
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub EnregistrementProductionAutreThread(ByVal NumCharge As Integer)

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                                                'pour les boucles FOR...NEXT
    Dim b As Integer                                                'pour les boucles FOR...NEXT
    Dim NumRedresseur As Integer                        'numéro d'un redresseur
    
    Dim MsgTracabilite As String
    
    Dim NumFicheProduction As String                   'numéro de fiche de production
    
    Dim ConnexionBDAnodisationSQL As ADODB.Connection
    Dim Enregistrement As ADODB.Recordset
    
   Dim FicheVideEtatsCharges As etatsCharges
    
    '--- affectation ---
    'EnregistrementProduction = ""
    
    'Call Log("ProchainNumFicheProduction  DEBUT")
    '--- recherche du prochain numéro de fiche de production ---
    NumFicheProduction = ProchainNumFicheProduction()
    'Call Log("ProchainNumFicheProduction  FIN")
                    
    If NumFicheProduction <> "" Then
    
        '--- ouverture de la connexion ---
        Set ConnexionBDAnodisationSQL = New ADODB.Connection
        With ConnexionBDAnodisationSQL
            .ConnectionString = PARAMETRES_CONNEXION_BD_ANODISATION_SQL
            .CursorLocation = adUseServer
            .Mode = adModeReadWrite
            .ConnectionTimeout = 2     'X secondes d'attente de connexion avant de lancer un message d'erreur
            .Open
        End With
        
        'Call Log("DETAILS DES CHARGES DE PRODUCTION  DEBUT")
        '--- extraction et enregistrement ---
        With TEtatsCharges(NumCharge)
    
            '****************************************************************************************************************
            '*                                                DETAILS DES CHARGES DE PRODUCTION
            '****************************************************************************************************************
            
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_CHARGES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
    
            '--- enregistrement des détails des charges ---
            For a = LBound(.TDetailsCharges()) To UBound(.TDetailsCharges())
                With .TDetailsCharges(a)
                    
                    If .NumCommandeInterne > 0 Then
                    
                        '--- enregistrement de la fiche ---
                        Enregistrement.AddNew
                        Enregistrement("NumCommandeInterne") = .NumCommandeInterne
                        Enregistrement("NbrReparations") = .TypeReparation
                        Enregistrement("DateEntreeEnLigne") = TEtatsCharges(NumCharge).DateEntreeEnLigne
                        Enregistrement("DateArriveeAuDechargement") = TEtatsCharges(NumCharge).DateArriveeAuDechargement
                        Enregistrement("NumBarre") = TEtatsCharges(NumCharge).NumBarre
                        Enregistrement("NumLigne") = a
                        Enregistrement("CodeClient") = .CodeClient
                        Enregistrement("NbrPieces") = .NbrPieces
                        Enregistrement("Designation") = .Designation
                        Enregistrement("NumLignesReferencesClient") = .NumLignesReferencesClient
                        Enregistrement("Matiere") = .Matiere
                        Enregistrement("NumGammeAnodisation") = TEtatsCharges(NumCharge).TGammesAnodisation.NumGamme
                        Enregistrement("RefGammeAnodisation") = TEtatsCharges(NumCharge).TGammesAnodisation.RefGamme
                        Enregistrement("TempsAnodisationTexte") = CTemps(TEtatsCharges(NumCharge).TempsTotalGammeRedresseur)
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        If TEtatsCharges(NumCharge).ChargePrioritaire = True Then
                            Enregistrement("ChargePrioritaire") = 1
                        Else
                            Enregistrement("ChargePrioritaire") = 0
                        End If
                        Enregistrement("AlarmesLigne") = TEtatsCharges(NumCharge).AlarmesLigne
                        'Enregistrement.Update
                    
                    Else
                        
                        '--- sortie directe si plus de n° de fiche détails de charge ---
                        Exit For
            
                    End If
                
                End With
            Next a
            
            Enregistrement.UpdateBatch
            
            Enregistrement.Close
            'Call Log("DETAILS DES CHARGES DE PRODUCTION  FIN")
            'Call Log("DETAILS DE LA GAMME D'ANODISATION DE PRODUCTION DEBUT")
        
            '****************************************************************************************************************
            '*                                      DETAILS DE LA GAMME D'ANODISATION DE PRODUCTION
            '****************************************************************************************************************
            
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_GAMMES_ANODISATION_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
            
            '--- enregistrement des détails de la gamme d'anodisation ---
            For a = LBound(.TGammesAnodisation.TDetailsGammesAnodisation()) To UBound(.TGammesAnodisation.TDetailsGammesAnodisation())
                
                With .TGammesAnodisation.TDetailsGammesAnodisation(a)
                    
                    If .NumZone <> 0 Then
                     
                        '--- enregistrement de la fiche ---
                        Enregistrement.AddNew
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        Enregistrement("NumLigne") = a
                        Enregistrement("NumZone") = .NumZone
                        Enregistrement("TempsAuPosteTexte") = .TempsAuPosteTexte
                        Enregistrement("TempsEgouttageTexte") = .TempsEgouttageTexte
                        Enregistrement("TempsAuPosteSecondes") = .TempsAuPosteSecondes
                        Enregistrement("TempsEgouttageSecondes") = .TempsEgouttageSecondes
                        Enregistrement("DecompteDuTempsAuPosteReelSecondes") = .DecompteDuTempsAuPosteReelSecondes
                        Enregistrement("NumPosteReel") = .NumPosteReel
                        
                        '--- affectation du numéro de redresseur ---
                        Select Case .NumPosteReel
                            Case POSTES.P_C13: NumRedresseur = REDRESSEURS.R_C13
                            Case POSTES.P_C14: NumRedresseur = REDRESSEURS.R_C14
                            Case POSTES.P_C15: NumRedresseur = REDRESSEURS.R_C15
                            Case POSTES.P_C16: NumRedresseur = REDRESSEURS.R_C16
                            Case Else
                        End Select
                        
                        '--- enregistrement ---
                        'Enregistrement.Update
                    
                    Else
                        
                        '--- sortie directe si plus de n° de fiche détails de charge ---
                        Exit For
            
                    End If
                
                End With
            Next a
            Enregistrement.UpdateBatch
            Enregistrement.Close
            'Call Log("DETAILS DE LA GAMME D'ANODISATION DE PRODUCTION FIN")
            
            'Call Log("TRACABILITE DES REDRESSEURS DEBUT")
            '****************************************************************************************************************
            '*                                                  TRACABILITE DES REDRESSEURS
            '****************************************************************************************************************
            If NumRedresseur > 0 Then                 'enregistrement de la production uniquement si passage dans un
                                                                          'des redresseurs
                Bidon = SauveTraçabiliteRedresseurs(NumCharge:=NumCharge, _
                                                                               NumFicheProduction:=NumFicheProduction, _
                                                                               DateEntreeEnLigne:=TEtatsCharges(NumCharge).DateEntreeEnLigne, _
                                                                               NumRedresseur:=NumRedresseur)
            
            End If

            'Call Log("TRACABILITE DES REDRESSEURS FIN")
            'Call Log("DETAILS DES PHASES DE PRODUCTION DEBUT")
            '****************************************************************************************************************
            '*                                       DETAILS DES PHASES DE PRODUCTION
            '****************************************************************************************************************
            
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_PHASES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
            
            '--- enregistrement des détails de la gamme d'anodisation ---
            For a = LBound(.TDetailsPhasesProduction()) To UBound(.TDetailsPhasesProduction())
                
                With .TDetailsPhasesProduction(a)
                    
                    '--- enregistrement de la fiche ---
                    Enregistrement.AddNew
                    Enregistrement("NumFicheProduction") = NumFicheProduction
                    Enregistrement("NumRedresseur") = NumRedresseur
                    Enregistrement("ModeUouI") = TEtatsCharges(NumCharge).ModeUouI
                    Enregistrement("NumPhase") = a
                    Enregistrement("TempsPhase") = .TempsPhase
                    Enregistrement("UPhase") = .UPhase
                    Enregistrement("IPhase") = .IPhase
                    'Enregistrement.Update
                    
                End With
            
            Next a
            Enregistrement.UpdateBatch
            Enregistrement.Close
            
            
            '****************************************************************************************************************
            '*                                                 DETAILS DES FICHES DE PRODUCTION
            '****************************************************************************************************************
        
            'Call Log("DETAILS DES PHASES DE PRODUCTION FIN")
            'Call Log("DETAILS DES FICHES DE PRODUCTION DEBUT")
            '--- ouverture de la table ---
            Set Enregistrement = New ADODB.Recordset
            With Enregistrement
                .CursorLocation = adUseServer
                .CursorType = adOpenStatic 'adOpenKeyset
                .LockType = adLockBatchOptimistic    'adLockOptimistic
                .Open TABLE_DETAILS_FICHES_PRODUCTION, ConnexionBDAnodisationSQL, , adCmdTable
            End With
        
            '--- enregistrement des détails des fiches de production ---
            For a = LBound(.TDetailsFichesProduction()) To UBound(.TDetailsFichesProduction())
                
                With .TDetailsFichesProduction(a)
                    
                    If .NumPoste <> 0 Then
                   
                        '--- enregistrement de la fiche ---
                        Enregistrement.AddNew
                        Enregistrement("NumFicheProduction") = NumFicheProduction
                        Enregistrement("NumLigne") = a
                        Enregistrement("NumPoste") = .NumPoste
                        Enregistrement("DateEntreePoste") = .DateEntreePoste
                        Enregistrement("DateSortiePoste") = .DateSortiePoste
                        Enregistrement("DateDebutEgouttage") = .DateDebutEgouttage
                        Enregistrement("DateFinEgouttage") = .DateFinEgouttage
                        Enregistrement("TemperatureEnEntree") = .TemperatureEnEntree
                        Enregistrement("TemperatureEnSortie") = .TemperatureEnSortie
                        Enregistrement("GrapheTemperature") = .GrapheTemperature
                        Enregistrement("URedresseur") = .URedresseur
                        Enregistrement("IRedresseur") = .IRedresseur
                        Enregistrement("GrapheRedresseur") = .GrapheRedresseur
                        Enregistrement("AnalyseurEnEntree") = .AnalyseurEnEntree
                        Enregistrement("AnalyseurEnSortie") = .AnalyseurEnSortie
                        Enregistrement("GrapheAnalyseur") = .GrapheAnalyseur
                        Enregistrement("AlarmesPoste") = .AlarmesPoste
                        'Enregistrement.Update
                    
                    Else
                   
                        '--- sortie directe si plus de n° de fiche détails de charge ---
                        Exit For
           
                    End If
           
                End With
            Next a
            Enregistrement.UpdateBatch
            Enregistrement.Close
        
            'Call Log("DETAILS DES FICHES DE PRODUCTION FIN")
        End With

    End If
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '****************************************************************************************************************
    '*                                        VIDAGE DE LA CHARGE DANS LE TABLEAU
    '****************************************************************************************************************
    TEtatsCharges(NumCharge) = FicheVideEtatsCharges
    
    '----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    '--- fermeture des enregistrements / connexions ---
    Select Case Enregistrement.State
        Case adStateClosed
        Case Else: Enregistrement.Close
    End Select
    ConnexionBDAnodisationSQL.Close
    
    '--- effacement des objets ---
    Set Enregistrement = Nothing
    Set ConnexionBDAnodisationSQL = Nothing
    
    Exit Sub

GestionErreurs:
    
    '--- valeur de retour ---
    'EnregistrementProduction = CStr(Err.Number)
    
    AfficheRenseignements ROUGE_0, "Erreur d'enregitrement en base: " & CStr(Err.Number) & vbCrLf
    
    '--- fermeture de l'enregistrement / connexion ---
    On Error Resume Next
    Enregistrement.Close
    Set Enregistrement = Nothing
    ConnexionBDAnodisationSQL.Close
    Set ConnexionBDAnodisationSQL = Nothing

End Sub



Private Function MakeTrue( _
                 ByRef bValue As Boolean) As Boolean
    MakeTrue = True
    bValue = True
End Function



Attribute VB_Name = "MQuestionsReponses"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DE GESTION DES QUESTIONS REPONSES
' Nom                    : MQuestionsReponses.bas
' Date de cr�ation : 08/11/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- constantes priv�es ---
Private Const LONGUEUR_MAXI_QUESTION As Integer = 100

'--- constantes publiques ---
Public Const NBR_PARAMETRES_POSSIBLES As Integer = 100
Public Const TABULATION_REPONSES As String = "     "
Public Const NOUVELLE_LIGNE As String = vbCrLf & TABULATION_REPONSES

Public Const MAUVAISE_FORMULATION As String = "Mauvaise formulation de la question"
Public Const PAS_DE_DISPOSITION_DU_PONT As String = "Vous ne disposez pas du contr�le du pont"
Public Const PAS_DE_DISPOSITION_DU_PONT_IA As String = "Le syst�me en CYCLIQUE ou IA ne dispose pas du contr�le du pont"
Public Const MOUVEMENTS_EN_COURS As String = "Mouvements en cours sur ce pont"
Public Const RISQUE_DE_COLLISION As String = "Risque de collision"
Public Const TRANSFERT_AUTOMATE_OK As String = "Transfert vers l'automate effectu� avec succ�s"
Public Const INCIDENT_TRANSFERT_AUTOMATE As String = "Incident d'�criture vers l'automate"

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extraire une question de la partie des dialogues
' Entr�es :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub ExtractionQuestion()
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- d�claration ---
    Dim CouleurReponse As Long
    Dim Caractere As String * 1
    Dim TexteADecoder As String, _
            Question As String, _
            Reponse As String
    Dim TEnsembleLignes As Variant

    '--- affectation ---
    TexteADecoder = Trim(Right(OccFSynoptique.RTBDialogues.Text, LONGUEUR_MAXI_QUESTION))
    TEnsembleLignes = Split(TexteADecoder, vbCrLf, , vbTextCompare)
        
    '--- extraction de la question ---
    If IsArray(TEnsembleLignes) = True Then
        
        '--- affectation de la question ---
        Question = TEnsembleLignes(UBound(TEnsembleLignes))
        
        '--- affectation de la r�ponse ---
        If Question = "" Then
            Reponse = vbCrLf
        Else
            Reponse = ReponseAUneQuestion(Question, CouleurReponse)
            If Reponse <> "" Then Reponse = Reponse & vbCrLf
        End If
        
        '--- affichage de la r�ponse ---
        Call OccFSynoptique.AfficheDialogues(CouleurReponse, Reponse)

    End If
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transmet l'attribution d'un num�ro de charge � un poste
' Entr�es :                  NumPoste -> Num�ro du poste concern�
'                               NumCharge -> Num�ro de charge
'                       CouleurReponse -> Couleur de la r�ponse
' Retours  : NumeroChargePoste -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function NumeroChargePoste(ByVal NumPoste As Variant, _
                                                            ByVal NumCharge As Variant, _
                                                            ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Reponse As String
                
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                    
        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                    
            '--- affectation ---
            Reponse = NOUVELLE_LIGNE & "LE NUMERO DE CHARGE AU POSTE " & _
                              TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                               " EST " & _
                              NumCharge & _
                              vbCrLf & TABULATION_REPONSES
                
            '--- envoi vers l'automate ---
            If EnvoiNumeroChargePoste(NumPoste, NumCharge) = OK Then
                CouleurReponse = COULEURS.BLEU_3
                Reponse = Reponse & TRANSFERT_AUTOMATE_OK
            Else
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = Reponse & INCIDENT_TRANSFERT_AUTOMATE
            End If
            
        Else
        
            '--- mauvaise formulation ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
        
        End If
    
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    NumeroChargePoste = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Pas de charge sur un poste (correspond � la mise � 0 du n� de charge)
' Entr�es :                 NumPoste -> Num�ro du poste concern�
'                              NumCharge -> Num�ro de charge
'                      CouleurReponse -> Couleur de la r�ponse
' Retours  : PasDeChargePoste -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PasDeChargePoste(ByVal NumPoste As Variant, _
                                                          ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Reponse As String
                
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE Then
                    
        '--- affectation ---
        Reponse = NOUVELLE_LIGNE & "PAS DE CHARGE AU POSTE " & _
                          TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                          vbCrLf & TABULATION_REPONSES
            
        '--- envoi vers l'automate ---
        If EnvoiNumeroChargePoste(NumPoste, 0) = OK Then
            CouleurReponse = COULEURS.BLEU_3
            Reponse = Reponse & TRANSFERT_AUTOMATE_OK
        Else
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = Reponse & INCIDENT_TRANSFERT_AUTOMATE
        End If
            
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    PasDeChargePoste = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transmet l'attribution d'un num�ro de charge � un pont
' Entr�es :                  NumPont -> Num�ro du pont concern�
'                             NumCharge -> Num�ro de charge
'                     CouleurReponse -> Couleur de la r�ponse
' Retours  : NumeroChargePont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function NumeroChargePont(ByVal NumPont As Variant, _
                                                          ByVal NumCharge As Variant, _
                                                          ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Reponse As String

    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        If NumCharge = 0 Or (NumCharge >= CHARGES.C_NUM_MINI And NumCharge <= CHARGES.C_NUM_MAXI) Then
                    
            '--- affectation ---
            Reponse = NOUVELLE_LIGNE & "LE NUMERO DE CHARGE SUR LE PONT " & _
                              NumPont & _
                               " EST " & _
                              NumCharge & _
                              vbCrLf & TABULATION_REPONSES
            
            '--- envoi vers l'automate ---
            If EnvoiNumeroChargePont(NumPont, NumCharge) = OK Then
                CouleurReponse = COULEURS.BLEU_3
                Reponse = Reponse & TRANSFERT_AUTOMATE_OK
            Else
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = Reponse & INCIDENT_TRANSFERT_AUTOMATE
            End If
                
        Else
        
            '--- mauvaise formulation ---
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
        
        End If
    
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    NumeroChargePont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Pas de charge sur un pont (correspond � la mise � 0 du n� de charge)
' Entr�es :                 NumPont -> Num�ro du pont concern�
'                    CouleurReponse -> Couleur de la r�ponse
' Retours  : PasDeChargePont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PasDeChargePont(ByVal NumPont As Variant, _
                                                        ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Reponse As String

    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
                    
        '--- affectation ---
        Reponse = NOUVELLE_LIGNE & "PAS DE CHARGE SUR LE PONT " & _
                          NumPont & _
                          vbCrLf & TABULATION_REPONSES
        
        '--- envoi vers l'automate ---
        If EnvoiNumeroChargePont(NumPont, 0) = OK Then
            CouleurReponse = COULEURS.BLEU_3
            Reponse = Reponse & TRANSFERT_AUTOMATE_OK
        Else
            CouleurReponse = COULEURS.ROUGE_3
            Reponse = Reponse & INCIDENT_TRANSFERT_AUTOMATE
        End If
                
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    PasDeChargePont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transmet le d�placement du pont au poste voulu
' Entr�es :                   NumPont -> Num�ro du pont concern�
'                                NumPoste -> Num�ro du poste souhait�
'                      CouleurReponse -> Couleur de la r�ponse
' Retours : NumeroChargePont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DeplacementPont(ByVal NumPont As Variant, _
                                                       ByVal NumPoste As Variant, _
                                                       ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    Dim Texte As String
    Texte = "DeplacementPont " & NumPont & ", NumPoste: " & NumPoste
    AfficheRenseignementsDebug CouleurReponse, Texte & vbCrLf
    
    
    '--- d�claration ---
    Dim a As Integer, _
           TypeCollision As Integer, _
           NumPontOppose As Integer, _
           NumPosteAssurantSecurite As Integer
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim CouleurReponseAntiCollision As Long
    Dim Reponse As String, _
            ReponseAntiCollision As String, _
            ReponseEnvoiCyclePont As String
    
    If NumPoste >= POSTES.P_CHGT_1 And NumPoste <= DERNIER_POSTE And _
       NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 Then
                    
        If TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
           TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then
        
            '*********************************************************************************************************
            '                               LES 2 PONTS SONT SOUS LE CONTROLE DE L'OPERATEUR
            '*********************************************************************************************************
            'il faut passer la commande dans le tableau des commandes op�rateur
            For a = LBound(TCommandesOperateur()) To UBound(TCommandesOperateur())
                With TCommandesOperateur(a)
                    If .TypeCycle = TYPES_CYCLES.TC_INCONNU Then           'remplissage de la commande si fiche vide
                        .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                        .NumPont = NumPont                                                        'num�ro du pont
                        .NumPosteDepart = TEtatsPonts(NumPont).PosteActuel  'sans int�r�t pour la commande, c'est
                                                                                                                   'juste pour mettre une valeur diff�rente
                                                                                                                   'de 0 pour l'anti-collision
                        .NumPosteArrivee = NumPoste                                         'num�ero du poste
                        .TempsEgouttageSecondes = 0                                         'temps d'�gouttage en secondes
                        Exit For
                    End If
                End With
            Next a
            
            '--- gestion de l'anti-collision ---
            ReponseAntiCollision = ControleAntiCollision(NumPont, _
                                                                                     TEtatsPonts(NumPont).PosteActuel, _
                                                                                     NumPoste, _
                                                                                     TypeCollision, _
                                                                                     NumPontOppose, _
                                                                                     NumPosteAssurantSecurite, _
                                                                                     CouleurReponseAntiCollision)
            
            '--- gestion de la r�ponse � l'anti-collision ---
            If NumPosteAssurantSecurite > 0 Or (NumPosteAssurantSecurite = 0 And TypeCollision = TYPES_COLLISION.AUCUN_RISQUE) Then
            
                '--- pas de risque de collision ---
                CouleurReponse = COULEURS.BLEU_3
                Reponse = NOUVELLE_LIGNE & "DEPLACEMENT DU PONT " & _
                                  NumPont & _
                                  " EN " & _
                                 TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                                 " MEMORISE"
            Else
                        
                '--- risque de collision ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = TABULATION_REPONSES & RISQUE_DE_COLLISION
                
            End If
                
        Else
            
            '*********************************************************************************************************
            '                                1 DES PONTS EST SOUS LE CONTROLE DE L'OPERATEUR
            '*********************************************************************************************************
            'la commande peut �tre lancer imm�diatement
            'celle-ci peut �tre annuler si il y a un risque de collision avec l'autre pont g�r�
            'par le moteur d'inf�rence
            
            '--- v�rification si l'op�rateur dispose du contr�le du pont ---
            If TEtatsPonts(NumPont).ControleParOperateur = True Then
                            
                '--- analyse si mouvements en cours ---
                If TEtatsPonts(NumPont).PtrEtActionEnCoursAPI.PtrAction = 0 Then
                
                    '--- gestion de l'anti-collision ---
                    ReponseAntiCollision = ControleAntiCollision(NumPont, _
                                                                                             TEtatsPonts(NumPont).PosteActuel, _
                                                                                             NumPoste, _
                                                                                             TypeCollision, _
                                                                                             NumPontOppose, _
                                                                                             NumPosteAssurantSecurite, _
                                                                                             CouleurReponseAntiCollision)
                    
                    If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
                        
                        '--- affectation ---
                        Reponse = NOUVELLE_LIGNE & "DEPLACEMENT DU PONT " & _
                                          NumPont & _
                                           " EN " & _
                                          TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & _
                                          vbCrLf & TABULATION_REPONSES
                    
                        '--- construction du cycle ---
                        Erase TUnCyclePont()
                        TUnCyclePont(1) = NumPoste
                        TUnCyclePont(2) = NUM_ACTION_FCY
                    
                        '--- lancement du d�placement ---
                        ReponseEnvoiCyclePont = EnvoiCyclePont(NumPont, TUnCyclePont)
                        Select Case ReponseEnvoiCyclePont
                            
                            Case OK
                                 '--- le cycle a �t� transf�r� avec succ�s, il faut remplir la fiche des param�tres ---
                                 With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                                    .NumPosteDepart = TEtatsPonts(NumPont).PosteActuel
                                    .NumPosteArrivee = NumPoste
                                    .TypeCycle = TYPES_CYCLES.TC_DEPLACEMENT_PONT
                                    .DelaiSupStabilisationChargeSecondes = 0
                                    .TempsEgouttageSecondes = 0
                                 End With
                                
                                '--- affectation de la r�ponse ---
                                CouleurReponse = COULEURS.BLEU_3
                                Reponse = Reponse & "D�placement au poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & " ACCEPTE"
                            
                            Case Else
                                '--- le d�placement a �t� refus� / affectation de la r�ponse ---
                                CouleurReponse = COULEURS.ROUGE_3
                                Reponse = Reponse & "D�placement au poste " & TEtatsPostes(NumPoste).DefinitionPoste.NomPoste & " REFUSE"
                                Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                        
                        End Select
                                        
                    Else
                        
                        '--- risque de collision ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = TABULATION_REPONSES & RISQUE_DE_COLLISION
                        
                    End If
                        
                Else
                
                    '--- des mouvements sont en cours ---
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MOUVEMENTS_EN_COURS
                
                End If
                
            Else
        
                '--- pas de disposition du pont ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = TABULATION_REPONSES & PAS_DE_DISPOSITION_DU_PONT & " " & NumPont
        
            End If
    
        End If
        
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    DeplacementPont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Transfert une charge d'un poste � un autre poste
' Entr�es :                     NumPosteDepart -> Num�ro du poste de d�part
'                                    NumPosteArrivee -> Num�ro du poste d'arriv�e
'                                                  NumPont -> Num�ro du pont souhait� pour le transfert
'                     TempsEgouttageSecondes -> Temps d'�gouttage en secondes
'                                     CouleurReponse -> Couleur de la r�ponse
' Retours :                      TransfertCharge -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function TransfertCharge(ByVal NumPosteDepart As Variant, _
                                                     ByVal NumPosteArrivee As Variant, _
                                                     ByVal NumPont As Variant, _
                                                     ByVal TempsEgouttageSecondes As Variant, _
                                                     ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
   
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer, _
           NumPontPremisse As Integer, _
           NumPontIAPremisse As Integer, _
           TypeCollision As Integer, _
           NumPontOppose As Integer, _
           NumPosteAssurantSecurite As Integer
    Dim TUnCyclePont(1 To NBR_LIGNES_CYCLES_PONTS) As Integer
    Dim CouleurReponseAntiCollision As Long, _
            TempsCycleSecondesPremisse As Long
    Dim Reponse As String, _
            ReponseAntiCollision As String, _
            ReponseExtraitPremisseDecodee As String, _
            ReponseEnvoiCyclePont As String
            
    If NumPosteDepart >= POSTES.P_CHGT_1 And NumPosteDepart <= DERNIER_POSTE And _
       NumPosteArrivee >= POSTES.P_CHGT_1 And NumPosteArrivee <= DERNIER_POSTE And _
       NumPont >= PONTS.P_1 And NumPont <= PONTS.P_2 Then
    
        If TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
           TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then
        
            '*********************************************************************************************************
            '                               LES 2 PONTS SONT SOUS LE CONTROLE DE L'OPERATEUR
            '*********************************************************************************************************
            'il faut passer la commande dans le tableau des commandes op�rateur
            For a = LBound(TCommandesOperateur()) To UBound(TCommandesOperateur())
                With TCommandesOperateur(a)
                    If .TypeCycle = TYPES_CYCLES.TC_INCONNU Then                 'remplissage de la commande si fiche vide
                        .TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE
                        .NumPont = NumPont                                                              'num�ro du pont
                        .NumPosteDepart = NumPosteDepart                                     'num�ro du poste de d�part
                        .NumPosteArrivee = NumPosteArrivee                                    'num�ro du poste d'arriv�e
                        .TempsEgouttageSecondes = TempsEgouttageSecondes      'temps d'�gouttage en secondes
                        Exit For
                    End If
                End With
            Next a
                        
            '--- affectation ---
            Reponse = NOUVELLE_LIGNE & "TRANSFERT DE LA CHARGE DE " & _
                              TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                              " EN " & _
                              TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                              " AVEC LE PONT " & NumPont & _
                              " MEMORISE"
            
        Else
    
            '*********************************************************************************************************
            '                                1 DES PONTS EST SOUS LE CONTROLE DE L'OPERATEUR
            '*********************************************************************************************************
            'la commande peut �tre lancer imm�diatement
            'celle-ci peut �tre annuler si il y a un risque de collision avec l'autre pont g�r�
            'par le moteur d'inf�rence
            
            '--- v�rification si l'op�rateur dispose du contr�le du pont ---
            If TEtatsPonts(NumPont).ControleParOperateur = True Then
        
                '--- analyse si mouvements en cours ---
                If TEtatsPonts(NumPont).PtrEtActionEnCoursAPI.PtrAction = 0 Then
                    
                    '--- gestion de l'anti-collision ---
                    ReponseAntiCollision = ControleAntiCollision(NumPont, _
                                                                                             NumPosteDepart, _
                                                                                             NumPosteArrivee, _
                                                                                             TypeCollision, _
                                                                                             NumPontOppose, _
                                                                                             NumPosteAssurantSecurite, _
                                                                                             CouleurReponseAntiCollision)
        
                    If TypeCollision = TYPES_COLLISION.AUCUN_RISQUE Then
    
                        '--- affectation ---
                        Reponse = NOUVELLE_LIGNE & "TRANSFERT DE LA CHARGE DE " & _
                                          TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                          " EN " & _
                                          TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                          " AVEC LE PONT " & NumPont & _
                                          vbCrLf & TABULATION_REPONSES
    
                        '--- effacement du tableau ---
                        Erase TUnCyclePont()
                    
                        '--- extraction de la pr�misse ---
                        ReponseExtraitPremisseDecodee = ExtraitPremisseDecodee(NumPosteDepart, _
                                                                                                                            NumPosteArrivee, _
                                                                                                                            NumPontPremisse, _
                                                                                                                            NumPontIAPremisse, _
                                                                                                                            TempsCycleSecondesPremisse, _
                                                                                                                            TUnCyclePont())
                
                        '--- v�rification de l'existence de la r�gle ---
                        If ReponseExtraitPremisseDecodee = OK Then
    
                            '--- insertion du temps d'�gouttage dans le cycle du pont ---
                            If IsNumeric(TempsEgouttageSecondes) = True Then
                                If CInt(TempsEgouttageSecondes) > 0 Then
                                    Bidon = InsertionTempsEgouttageDansCyclePont(TempsEgouttageSecondes, TUnCyclePont())
                                End If
                            End If
                            
                            '--- lancement du transfert ---
                            ReponseEnvoiCyclePont = EnvoiCyclePont(NumPont, TUnCyclePont())
                            Select Case ReponseEnvoiCyclePont
                                
                                Case OK
                                     '--- le cycle a �t� transf�r� avec succ�s, il faut remplir la fiche des param�tres ---
                                     With TEtatsPonts(NumPont).TParametresCyclesPonts(CYCLES.C_ACTUEL)
                                        .NumPosteDepart = NumPosteDepart
                                        .NumPosteArrivee = NumPosteArrivee
                                        .TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE
                                        .DelaiSupStabilisationChargeSecondes = 0
                                        .TempsEgouttageSecondes = TempsEgouttageSecondes
                                     End With
                                    
                                    '--- affectation de la r�ponse ---
                                    CouleurReponse = COULEURS.BLEU_3
                                    Reponse = Reponse & "Transfert de la charge de " & TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                                      " en " & TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                                      " avec le pont " & NumPont & _
                                                     IIf(TempsEgouttageSecondes = 0, "", ", �gouttage " & TempsEgouttageSecondes & " secondes") & _
                                                      " ACCEPTE"
                                Case Else
                                    '--- le transfert a �t� refus� / affectation de la r�ponse ---
                                    CouleurReponse = COULEURS.ROUGE_3
                                    Reponse = Reponse & "Transfert de la charge de " & TEtatsPostes(NumPosteDepart).DefinitionPoste.NomPoste & _
                                                      " en " & TEtatsPostes(NumPosteArrivee).DefinitionPoste.NomPoste & _
                                                      " avec le pont " & NumPont & _
                                                      IIf(TempsEgouttageSecondes = 0, "", ", �gouttage " & TempsEgouttageSecondes & " secondes") & _
                                                      " REFUSE"
                                    Reponse = Reponse & vbCrLf & TABULATION_REPONSES & ReponseEnvoiCyclePont
                            
                            End Select
    
                        Else
            
                            '--- mauvaise formulation ---
                            CouleurReponse = COULEURS.ROUGE_3
                            Reponse = TABULATION_REPONSES & ReponseExtraitPremisseDecodee
            
                        End If
    
                    Else
                        
                        '--- risque de collision ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = TABULATION_REPONSES & RISQUE_DE_COLLISION
                        
                    End If
                            
                Else
                
                    '--- des mouvements sont en cours ---
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MOUVEMENTS_EN_COURS
                
                End If
                
            Else
    
                '--- pas de disposition du pont ---
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = TABULATION_REPONSES & PAS_DE_DISPOSITION_DU_PONT & " " & NumPont
    
            End If
    
        End If
    
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
    
    End If

    '--- valeur de retour ---
    TransfertCharge = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet de prendre le contr�le d'un pont pour effectuer son pilotage � la demande (uniquement en auto)
' Entr�es :              NumPont -> Num�ro du pont concern�
'                 CouleurReponse -> Couleur de la r�ponse
' Retours  :       ControlePont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControlePont(ByVal NumPont As Variant, _
                                                ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Reponse As String
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
        
        '--- affectation ---
        Reponse = NOUVELLE_LIGNE & "CONTROLE DU PONT " & _
                           IIf(NumPont = PONTS.P_1, "1", "2") & _
                           vbCrLf & TABULATION_REPONSES
        
        '--- r�ponse ---
        With TEtatsPonts(NumPont)
            If .ModePont = MODES_PONTS.M_AUTOMATIQUE Then
                If .ControleParOperateur = True Then
                    CouleurReponse = COULEURS.ORANGE_3
                    Reponse = Reponse & "Le pont " & NumPont & " est d�j� sous votre contr�le"
                Else
                    .ControleParOperateur = True
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = Reponse & "Contr�le du pont " & NumPont & " autoris� (apr�s la fin de la s�quence en cours)"
                End If
            Else
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = Reponse & "Contr�le du pont " & NumPont & " impossible car celui-ci n'est pas en automatique"
            End If
        End With
    
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
        
    End If

    '--- valeur de retour ---
    ControlePont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Permet de rendre le contr�le d'un pont au syst�me
' Entr�es :               NumPont -> Num�ro du pont concern�
'                  CouleurReponse -> Couleur de la r�ponse
' Retours : RecuperationPont -> Message � retourner comme r�ponse
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function RepriseSystemePont(ByVal NumPont As Integer, _
                                                             ByRef CouleurReponse As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim Reponse As String
    
    If NumPont = PONTS.P_1 Or NumPont = PONTS.P_2 Then
        
        '--- affectation ---
        Reponse = NOUVELLE_LIGNE & "REPRISE PAR LE SYSTEME DU PONT " & _
                          NumPont & vbCrLf & TABULATION_REPONSES
                                               
        '--- r�ponse ---
        With TEtatsPonts(NumPont)
            If .ModePont = MODES_PONTS.M_AUTOMATIQUE Then
                If .ControleParOperateur = False Then
                    CouleurReponse = COULEURS.ORANGE_3
                    Reponse = Reponse & "Le pont " & NumPont & " est d�j� g�r� par le syst�me"
                Else
                    .ControleParOperateur = False
                    CouleurReponse = COULEURS.BLEU_3
                    Reponse = Reponse & "Reprise par le syst�me du pont " & NumPont & " effectu�e"
                End If
            Else
                CouleurReponse = COULEURS.ROUGE_3
                Reponse = Reponse & "R�cup�ration du pont " & NumPont & " impossible car celui-ci n'est pas en automatique"
            End If
        End With
    
    Else
        
        '--- mauvaise formulation ---
        CouleurReponse = COULEURS.ROUGE_3
        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
        
    End If

    '--- valeur de retour ---
    RepriseSystemePont = Reponse

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : R�ponse � une question pos�e par l'op�rateur
' Entr�es :                        Question -> Question pos�e par l'op�rateur et � d�coder
' Retours : ReponseAUneQuestion -> Confirmation de la question � afficher dans la zone des dialogues
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ReponseAUneQuestion(ByVal Question As String, _
                                                                ByRef CouleurReponse As Long) As String

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim PontCommeParametre As Boolean, _
            PosteCommeParametre As Boolean
    Dim a As Integer, _
           b As Integer, _
           NbrParametres As Integer, _
           NumParametre As Integer, _
           NumPont As Integer, _
           NumPoste As Integer, _
           NumCharge As Integer, _
           TempsEgouttageSecondes As Integer
    Dim CodageQuestion As String, _
            Reponse As String
    Dim TMotsQuestion As Variant, _
            TParametres(1 To NBR_PARAMETRES_POSSIBLES) As Variant
    Static TMemQuestions(1 To 10) As String                                 'tableau m�morisant toutes les questions
    
    '--- affectation ---
    CouleurReponse = COULEURS.BLEU_3
    
    '--- v�rification de la question ---
    Question = Trim(UCase(Question))
    If Question = "" Then
        
        '--- sortie de la fonction car la question est inutile ---
        ReponseAUneQuestion = ""
        Exit Function
    
    Else
        
        '--- m�morisation des questions pour un rappel ult�rieur ---
        If Question <> "R" And Question <> "RAPPEL" Then                'ne pas m�moriser la commande RAPPEL
            For a = UBound(TMemQuestions()) To (LBound(TMemQuestions()) + 1) Step -1
                TMemQuestions(a) = TMemQuestions(a - 1)
            Next a
            TMemQuestions(1) = Question
        End If
    
    End If
    
    '--- s�paration des mots ---
    TMotsQuestion = Split(Question, UN_ESPACE)
    
    '--- recherche d'une �quivalence � la question compl�te ---
    If LBound(TMotsQuestion) = 0 And UBound(TMotsQuestion) = 0 Then
        Select Case TMotsQuestion(0)
            Case "A"
                '--- annulation de toutes les commandes ---
                TMotsQuestion = Split("ANNULATION DE LA TRACABILITE DES COMMANDES", UN_ESPACE)
            Case "I"
                '--- initialisation des charges pour la journ�e portes ouvertes ---
                TMotsQuestion = Split("INITIALISATION POUR LA JOURNEE PORTES OUVERTES", UN_ESPACE)
            Case "ED", "CLS"
                '--- effacement des dialogues ---
                TMotsQuestion = Split("EFFACEMENT DES DIALOGUES", UN_ESPACE)
            Case "CP1"
                '--- contr�le du pont 1 ---
                TMotsQuestion = Split("CONTROLE PONT 1", UN_ESPACE)
            Case "CP2"
                '--- contr�le du pont 2 ---
                TMotsQuestion = Split("CONTROLE PONT 2", UN_ESPACE)
            Case "CP1-2"
                '--- contr�le du pont 1 et 2 ---
                TMotsQuestion = Split("CONTROLE DU PONT 1 ET 2", UN_ESPACE)
            Case "CP2-1"
                '--- contr�le du pont 2 et 1 ---
                TMotsQuestion = Split("CONTROLE DU PONT 2 ET 1", UN_ESPACE)
            Case "RSP1"
                '--- reprise par le syst�me du pont 1 ---
                TMotsQuestion = Split("REPRISE SYSTEME PONT 1", UN_ESPACE)
            Case "RSP2"
                '--- reprise par le syst�me du pont 2 ---
                TMotsQuestion = Split("REPRISE SYSTEME PONT 2", UN_ESPACE)
            Case "RSP1-2"
                '--- reprise par le syst�me du pont 1 et 2 ---
                TMotsQuestion = Split("REPRISE SYSTEME DU PONT 1 ET 2", UN_ESPACE)
            Case "RSP2-1"
                '--- reprise par le syst�me du pont 2 et 1 ---
                TMotsQuestion = Split("REPRISE SYSTEME DU PONT 2 ET 1", UN_ESPACE)
            Case Else
        End Select
    End If
    
    '--- affectation ---
    NumParametre = 1
    CodageQuestion = ""
    
    '--- d�composition de la question ---
    For a = LBound(TMotsQuestion) To UBound(TMotsQuestion)
            
        Select Case TMotsQuestion(a)
            
            Case "", " ", "=", "LE", "LA", "LES", "DE", "DU", "SUR", "PAR", "EN", "AU", "AVEC", "EST", "ET", "DES"                      'valeurs � �liminer
            
            Case "P1": TParametres(NumParametre) = PONTS.P_1: PontCommeParametre = True: Inc NumParametre
            Case "P2": TParametres(NumParametre) = PONTS.P_2: PontCommeParametre = True: Inc NumParametre
            
            Case "CHGT1": TParametres(NumParametre) = POSTES.P_CHGT_1: PosteCommeParametre = True: Inc NumParametre
            Case "CHGT2": TParametres(NumParametre) = POSTES.P_CHGT_2: PosteCommeParametre = True: Inc NumParametre
            'Case "CHGT3": TParametres(NumParametre) = POSTES.P_CHGT_3: PosteCommeParametre = True: Inc NumParametre
            'Case "CHGT4": TParametres(NumParametre) = POSTES.P_CHGT_2: PosteCommeParametre = True: Inc NumParametre
            
            Case "C00": TParametres(NumParametre) = POSTES.P_C00: PosteCommeParametre = True: Inc NumParametre
            Case "DEC": TParametres(NumParametre) = POSTES.P_DEC: PosteCommeParametre = True: Inc NumParametre
            Case "SAT": TParametres(NumParametre) = POSTES.P_SAT: PosteCommeParametre = True: Inc NumParametre
            Case "C02": TParametres(NumParametre) = POSTES.P_C02: PosteCommeParametre = True: Inc NumParametre
            Case "C03": TParametres(NumParametre) = POSTES.P_C03: PosteCommeParametre = True: Inc NumParametre
            Case "C04": TParametres(NumParametre) = POSTES.P_C04: PosteCommeParametre = True: Inc NumParametre
            Case "C05": TParametres(NumParametre) = POSTES.P_C05: PosteCommeParametre = True: Inc NumParametre
            Case "C06": TParametres(NumParametre) = POSTES.P_C06: PosteCommeParametre = True: Inc NumParametre
            Case "C07": TParametres(NumParametre) = POSTES.P_C07: PosteCommeParametre = True: Inc NumParametre
            Case "C08": TParametres(NumParametre) = POSTES.P_C08: PosteCommeParametre = True: Inc NumParametre
            Case "C09": TParametres(NumParametre) = POSTES.P_C09: PosteCommeParametre = True: Inc NumParametre
            Case "C10": TParametres(NumParametre) = POSTES.P_C10: PosteCommeParametre = True: Inc NumParametre
            Case "C11": TParametres(NumParametre) = POSTES.P_C11: PosteCommeParametre = True: Inc NumParametre
            Case "C12": TParametres(NumParametre) = POSTES.P_C12: PosteCommeParametre = True: Inc NumParametre
            Case "C13": TParametres(NumParametre) = POSTES.P_C13: PosteCommeParametre = True: Inc NumParametre
            Case "C14": TParametres(NumParametre) = POSTES.P_C14: PosteCommeParametre = True: Inc NumParametre
            Case "C15": TParametres(NumParametre) = POSTES.P_C15: PosteCommeParametre = True: Inc NumParametre
            Case "C16": TParametres(NumParametre) = POSTES.P_C16: PosteCommeParametre = True: Inc NumParametre
            Case "C17": TParametres(NumParametre) = POSTES.P_C17: PosteCommeParametre = True: Inc NumParametre
            Case "C18": TParametres(NumParametre) = POSTES.P_C18: PosteCommeParametre = True: Inc NumParametre
            Case "C19": TParametres(NumParametre) = POSTES.P_C19: PosteCommeParametre = True: Inc NumParametre
            Case "C20": TParametres(NumParametre) = POSTES.P_C20: PosteCommeParametre = True: Inc NumParametre
            Case "C21": TParametres(NumParametre) = POSTES.P_C21: PosteCommeParametre = True: Inc NumParametre
            Case "C22": TParametres(NumParametre) = POSTES.P_C22: PosteCommeParametre = True: Inc NumParametre
            Case "C23": TParametres(NumParametre) = POSTES.P_C23: PosteCommeParametre = True: Inc NumParametre
            Case "C24": TParametres(NumParametre) = POSTES.P_C24: PosteCommeParametre = True: Inc NumParametre
            Case "C25": TParametres(NumParametre) = POSTES.P_C25: PosteCommeParametre = True: Inc NumParametre
            Case "C26": TParametres(NumParametre) = POSTES.P_C26: PosteCommeParametre = True: Inc NumParametre
            Case "C27": TParametres(NumParametre) = POSTES.P_C27: PosteCommeParametre = True: Inc NumParametre
            Case "C28": TParametres(NumParametre) = POSTES.P_C28: PosteCommeParametre = True: Inc NumParametre
            Case "C29": TParametres(NumParametre) = POSTES.P_C29: PosteCommeParametre = True: Inc NumParametre
            Case "C30": TParametres(NumParametre) = POSTES.P_C30: PosteCommeParametre = True: Inc NumParametre
            Case "C31": TParametres(NumParametre) = POSTES.P_C31: PosteCommeParametre = True: Inc NumParametre
            Case "C32": TParametres(NumParametre) = POSTES.P_C32: PosteCommeParametre = True: Inc NumParametre
            Case "C33": TParametres(NumParametre) = POSTES.P_C33: PosteCommeParametre = True: Inc NumParametre
            Case "C34": TParametres(NumParametre) = POSTES.P_C34: PosteCommeParametre = True: Inc NumParametre
            Case "C35": TParametres(NumParametre) = POSTES.P_C35: PosteCommeParametre = True: Inc NumParametre
            
            Case "D1": TParametres(NumParametre) = POSTES.P_D1: PosteCommeParametre = True: Inc NumParametre
            Case "D2": TParametres(NumParametre) = POSTES.P_D2: PosteCommeParametre = True: Inc NumParametre
            
            ' Case "C36": TParametres(NumParametre) = POSTES.P_C36: PosteCommeParametre = True: Inc NumParametre
             Case "C37": TParametres(NumParametre) = POSTES.P_C37: PosteCommeParametre = True: Inc NumParametre
            Case "C38": TParametres(NumParametre) = POSTES.P_C38: PosteCommeParametre = True: Inc NumParametre
            
            Case Else
                If IsNumeric(TMotsQuestion(a)) = True Then
                    
                    '--- si nombre alors param�tre ---
                    TParametres(NumParametre) = TMotsQuestion(a)
                    Inc NumParametre
                
                Else
                    
                    '--- v�rification du type de param�tre ---
                    If TMotsQuestion(a) = "PONT" Then PontCommeParametre = True
                    If TMotsQuestion(a) = "POSTE" Then PosteCommeParametre = True
                    
                    '--- ajouter la premi�re lettre au codage de la question ---
                    CodageQuestion = CodageQuestion & Left(TMotsQuestion(a), 1)
                
                End If
        
        End Select
    
    Next a
    
    '--- affectation du nombre de param�tres ---
    NbrParametres = Pred(NumParametre)
    If NbrParametres < 0 Then NbrParametres = 0
    
    With OccFSynoptique
    
        '--- analyse de la question en fonction de son codage ---
        Select Case CodageQuestion
            
            Case "ATC"
                '****************************************************************************************************
                '                                        Annulation de la tra�abilit� des commandes
                '****************************************************************************************************
                '--- effacement du tableau des commandes ---
                Erase TCommandesOperateur()
                
                '--- affectation de la r�ponse ---
                CouleurReponse = COULEURS.BLEU_3
                Reponse = NOUVELLE_LIGNE & "ANNULATION DE LA TRACABILITE DES COMMANDES"
                
            Case "B"
                '****************************************************************************************************
                '                                                                  Bonjour
                '****************************************************************************************************
                CouleurReponse = COULEURS.BLEU_3
                Reponse = NOUVELLE_LIGNE & "Bonjour"
            
            Case "ED"
                '****************************************************************************************************
                '                                          Effacement de la zone des dialogues
                '****************************************************************************************************
                With OccFSynoptique.RTBDialogues
                    .Text = ""
                    .Refresh
                    If .Visible = True Then .SetFocus
                End With

            Case "NC", "NCP"
                '****************************************************************************************************
                '                                           Num�ro de charge au poste et pont
                '****************************************************************************************************
                'Phrase compl�te : LE NUMERO DE CHARGE AU POSTE A1 EST 10
                '         Abr�viation : N C A1 10
                'Phrase compl�te : LE NUMERO DE CHARGE SUR LE PONT 1 EST 10
                '         Abr�viation : N C P1 10
                If PosteCommeParametre = True And NbrParametres = 2 Then
                    Reponse = NumeroChargePoste(NumPoste:=TParametres(1), _
                                                                         NumCharge:=TParametres(2), _
                                                                         CouleurReponse:=CouleurReponse)
                ElseIf PontCommeParametre = True And NbrParametres = 2 Then
                    Reponse = NumeroChargePont(NumPont:=TParametres(1), _
                                                                       NumCharge:=TParametres(2), _
                                                                       CouleurReponse:=CouleurReponse)
                Else
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
            
            Case "PC", "PCP"
                '****************************************************************************************************
                '                                             Pas de charge � un poste ou pont
                '****************************************************************************************************
                'Phrase compl�te : PAS DE CHARGE AU POSTE A1
                '         Abr�viation : P C A1
                'Phrase compl�te : PAS DE CHARGE SUR LE PONT 1
                '         Abr�viation : P C P1
                If PosteCommeParametre = True And NbrParametres = 1 Then
                    Reponse = PasDeChargePoste(NumPoste:=TParametres(1), _
                                                                         CouleurReponse:=CouleurReponse)
                ElseIf PontCommeParametre = True And NbrParametres = 1 Then
                    Reponse = PasDeChargePont(NumPont:=TParametres(1), _
                                                                     CouleurReponse:=CouleurReponse)
                Else
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
            
            Case "CP"
                '****************************************************************************************************
                '                                                   Contr�le du pont 1 et 2
                '****************************************************************************************************
                'Phrase compl�te : CONTROLE DU PONT 1
                '         Abr�viation : CP1
                'Phrase compl�te : CONTROLE DU PONT 2
                '         Abr�viation : CP2
                'Phrase compl�te : CONTROLE DU PONT 1 ET 2
                '         Abr�viation : CP1-2
                'Phrase compl�te : CONTROLE DU PONT 2 ET 1
                '         Abr�viation : CP2-1
                If PontCommeParametre = True And NbrParametres = 1 Then
                    Reponse = ControlePont(NumPont:=TParametres(1), _
                                                            CouleurReponse:=CouleurReponse)
                ElseIf PontCommeParametre = True And NbrParametres = 2 Then
                    Reponse = ControlePont(NumPont:=TParametres(1), _
                                                            CouleurReponse:=CouleurReponse) & _
                                      ControlePont(NumPont:=TParametres(2), _
                                                            CouleurReponse:=CouleurReponse)
                Else
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
            
            Case "RSP"
                '****************************************************************************************************
                '                                             Reprise syst�me du pont 1 et 2
                '****************************************************************************************************
                'Phrase compl�te : REPRISE SYSTEME DU PONT 1
                '         Abr�viation : RSP1
                'Phrase compl�te : REPRISE SYSTEME DU PONT 2
                '         Abr�viation : RSP2
                'Phrase compl�te : REPRISE SYSTEME DU PONT 1 ET 2
                '         Abr�viation : RSP1-2
                'Phrase compl�te : REPRISE SYSTEME DU PONT 2 ET 1
                '         Abr�viation : RSP2-1
                If PontCommeParametre = True And NbrParametres = 1 Then
                    Reponse = RepriseSystemePont(NumPont:=TParametres(1), _
                                                                          CouleurReponse:=CouleurReponse)
                ElseIf PontCommeParametre = True And NbrParametres = 2 Then
                    Reponse = RepriseSystemePont(NumPont:=TParametres(1), _
                                                                          CouleurReponse:=CouleurReponse) & _
                                      RepriseSystemePont(NumPont:=TParametres(2), _
                                                                          CouleurReponse:=CouleurReponse)
                
                Else
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
            
            Case "D", "DP"
                '****************************************************************************************************
                '                                                  D�placement d'un pont
                '****************************************************************************************************
                'Phrase compl�te : DEPLACEMENT DU PONT 1 EN C1
                '         Abr�viation : D P1 C1
                'Phrase compl�te : DEPLACEMENT DU PONT 2 EN D1
                '         Abr�viation : D P2 D1
                If PontCommeParametre = True And NbrParametres = 2 Then
                    Reponse = DeplacementPont(NumPont:=TParametres(1), _
                                                                    NumPoste:=TParametres(2), _
                                                                    CouleurReponse:=CouleurReponse)
                Else
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
            
            Case "TC", "TCE", "TCP", "TCPE"
                '****************************************************************************************************
                '                                    Transfert d'une charge d'un poste � un autre
                '****************************************************************************************************
                'Phrase compl�te : TRANSFERT DE LA CHARGE DE C3 EN A1
                '         Abr�viation : T C C3 A1
                'Phrase compl�te : TRANSFERT DE LA CHARGE DE C3 EN A1 AVEC LE PONT 1
                '         Abr�viation : T C C3 A1 P1
                'ou avec �gouttage
                'Phrase compl�te : TRANSFERT DE LA CHARGE DE C3 EN A1 EGOUTTAGE = 10
                '         Abr�viation : T C C3 A1 E 10
                'Phrase compl�te : TRANSFERT DE LA CHARGE DE C3 EN A1 AVEC LE PONT 1 EGOUTTAGE = 10
                '         Abr�viation : T C C3 A1 P1 E 10
                If PosteCommeParametre = True And PontCommeParametre = False And NbrParametres = 2 Then
                    '--- Abr�viation : T C C3 A1 ---
                    Reponse = TransfertCharge(NumPosteDepart:=TParametres(1), _
                                                                 NumPosteArrivee:=TParametres(2), _
                                                                 NumPont:=TPremisses(TParametres(1), TParametres(2)).NumPont, _
                                                                 TempsEgouttageSecondes:=0, _
                                                                 CouleurReponse:=CouleurReponse)
                ElseIf PosteCommeParametre = True And PontCommeParametre = True And NbrParametres = 3 Then
                    '--- Abr�viation : T C C3 A1 P1 ---
                    Reponse = TransfertCharge(NumPosteDepart:=TParametres(1), _
                                                                 NumPosteArrivee:=TParametres(2), _
                                                                 NumPont:=TParametres(3), _
                                                                 TempsEgouttageSecondes:=0, _
                                                                 CouleurReponse:=CouleurReponse)
                ElseIf PosteCommeParametre = True And PontCommeParametre = False And NbrParametres = 3 Then
                    '--- Abr�viation : T C C3 A1 E 10 ---
                    Reponse = TransfertCharge(NumPosteDepart:=TParametres(1), _
                                                                 NumPosteArrivee:=TParametres(2), _
                                                                 NumPont:=TPremisses(TParametres(1), TParametres(2)).NumPont, _
                                                                 TempsEgouttageSecondes:=TParametres(3), _
                                                                 CouleurReponse:=CouleurReponse)
                ElseIf PosteCommeParametre = True And PontCommeParametre = True And NbrParametres = 4 Then
                    '--- Abr�viation : T C C3 A1 P1 E 10 ---
                    Reponse = TransfertCharge(NumPosteDepart:=TParametres(1), _
                                                                 NumPosteArrivee:=TParametres(2), _
                                                                 NumPont:=TParametres(3), _
                                                                 TempsEgouttageSecondes:=TParametres(4), _
                                                                 CouleurReponse:=CouleurReponse)
                Else
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
            
            Case "TME"
                '****************************************************************************************************
                '            Transferts multiples d'une charge d'un poste � un autre, puis � un autre, etc ...
                '****************************************************************************************************
                'Phrase compl�te : TRANSFERTS MULTILPES D'UNE CHARGE
                '         Abr�viation : T M C3 A1 A4 A5 A6 P1 E 10
                If TEtatsPonts(PONTS.P_1).ControleParOperateur = True And _
                   TEtatsPonts(PONTS.P_2).ControleParOperateur = True Then
                    If PosteCommeParametre = True And PontCommeParametre = True And NbrParametres >= 4 Then
                        
                        '--- extraction du n� de pont et temps d'�gouttage ---
                        TempsEgouttageSecondes = TParametres(NbrParametres)   'le temps d'�gouttage est
                                                                                                                          'le dernier param�tre
                        NumPont = TParametres(NbrParametres - 1)                           'le num�ro du pont est l'avant
                                                                                                                          'dernier param�tre
                        
                        '--- affectation de toutes les commandes ---
                        For a = 1 To NbrParametres - 3
                            For b = LBound(TCommandesOperateur()) To UBound(TCommandesOperateur())
                                With TCommandesOperateur(b)
                                    If .TypeCycle = TYPES_CYCLES.TC_INCONNU Then                 'remplissage de la commande si fiche vide
                                        .TypeCycle = TYPES_CYCLES.TC_TRANSFERT_CHARGE
                                        .NumPont = NumPont                                                              'num�ro du pont
                                        .NumPosteDepart = TParametres(a)                                        'num�ro du poste de d�part
                                        .NumPosteArrivee = TParametres(a + 1)                                 'num�ro du poste d'arriv�e
                                        .TempsEgouttageSecondes = TempsEgouttageSecondes      'temps d'�gouttage en secondes
                                        Exit For
                                    End If
                                End With
                            Next b
                        Next a

                        '--- affectation de la r�ponse ---
                        CouleurReponse = COULEURS.BLEU_3
                        Reponse = NOUVELLE_LIGNE & "TRANSFERTS MULTILPES D'UNE CHARGE MEMORISES"
                    
                    Else
                        
                        '--- affectation de la r�ponse ---
                        CouleurReponse = COULEURS.ROUGE_3
                        Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                    
                    End If
                
                Else
                    
                    '--- affectation de la r�ponse ---
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & "Vous ne disposez pas du contr�le des 2 ponts pour lancer cette commande"
                    
                End If
            
            Case "R"
                '****************************************************************************************************
                '                                 Rappel de la m�moire des 10 derni�res commandes
                '****************************************************************************************************
                Reponse = TABULATION_REPONSES & "RAPPEL" & vbCrLf
                For a = UBound(TMemQuestions()) To LBound(TMemQuestions()) Step -1
                    If Trim(TMemQuestions(a)) <> "" Then
                        Reponse = Reponse & Trim(TMemQuestions(a))
                        If a <> LBound(TMemQuestions()) Then
                            If Trim(TMemQuestions(Pred(a))) <> "" Then
                                Reponse = Reponse & vbCrLf    'ne mettre un retour chariot uniquement si une commande suit
                            End If
                        End If
                    End If
                Next a
            
            Case Else
                '--- mauvaise formulation ---
                If Question <> "" Then
                    CouleurReponse = COULEURS.ROUGE_3
                    Reponse = TABULATION_REPONSES & MAUVAISE_FORMULATION
                End If
                    
        End Select
        
        '--- couleur par d�faut ---
    
    End With
    
    '--- valeur de retour ---
    ReponseAUneQuestion = Reponse

End Function


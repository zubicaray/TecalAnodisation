Attribute VB_Name = "MImpressions"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE GERANT LES IMPRESSIONS
' Nom                    : MImpressions.bas
' Date de cr�ation : 04/06/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : G�re les impressions en fonction de la fenetre active
' Entr�es : EtatSouhaite -> Fonction de l'�num�ration TYPE_IMPRESSIONS
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub Impressions(ByVal EtatSouhaite As TYPES_IMPRESSIONS)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const TITRE_MESSAGES As String = "Impressions"

    '--- d�claration ---
    Dim NumfenetreActive As Long

    '--- affectation ---
    NumfenetreActive = OccFPrincipale.ActiveForm.NumFenetre

    '--- appel de la fen�tre ---
    Select Case NumfenetreActive

        Case FENETRES.F_SYNOPTIQUE
            '--- op�rations � effectuer ---
            Select Case EtatSouhaite
                Case TYPES_IMPRESSIONS.TI_APERCU_AVANT_IMPRESSION, TYPES_IMPRESSIONS.TI_IMPRIMER
                Case TYPES_IMPRESSIONS.TI_IMPRIMER_FENETRE_ACTIVE
                    '--- impression de la fen�tre active ---
                    Printer.Orientation = vbPRORLandscape
                    OccFPrincipale.ActiveForm.PrintForm
                    Printer.Orientation = vbPRORPortrait
                
                Case Else
            End Select
        
        Case FENETRES.F_GAMMES_ANODISATION
            '--- op�rations � effectuer ---
            Select Case EtatSouhaite
                Case TYPES_IMPRESSIONS.TI_APERCU_AVANT_IMPRESSION, TYPES_IMPRESSIONS.TI_IMPRIMER

                    '--- contr�les et appel de l'�cran des impressions ---
                    If PossibiliteImpression = False Then Exit Sub
                    OptionImpressionChoisie = ChoixImpression(NumfenetreActive, "La gamme d'anodisation en cours", "Toutes les gammes d'anodisation")

                    '--- affectation des crit�res d'impression ---
                    Select Case OptionImpressionChoisie
                        Case 1: TCriteresImpression(1) = OccFPrincipale.ActiveForm.TBNumGamme.Text
                        Case 2
                        Case Else
                    End Select
                    
                    '--- lancement de l'aper�u ou de l'impression ---
                    Call OccFPrincipale.ActiveForm.SourisEnAttente(True)
                    If EtatSouhaite = TYPES_IMPRESSIONS.TI_APERCU_AVANT_IMPRESSION Then
                        Select Case OptionImpressionChoisie
                            Case 1
                                DRGammesAnodisation1.Show vbModal, OccFPrincipale
                                DRGammesAnodisation2.Show vbModal, OccFPrincipale
                            Case 2
                                DRGammesAnodisation3.Show vbModal, OccFPrincipale
                            Case Else
                        End Select
                    Else
                        Select Case OptionImpressionChoisie
                            Case 1
                                DRGammesAnodisation1.PrintReport
                                DRGammesAnodisation2.PrintReport
                            Case 2
                                DRGammesAnodisation3.PrintReport
                            Case Else
                        End Select
                    End If
                    Call OccFPrincipale.ActiveForm.SourisEnAttente(False)

                Case TYPES_IMPRESSIONS.TI_IMPRIMER_FENETRE_ACTIVE
                    '--- impression de la fen�tre active ---
                    Printer.Orientation = vbPRORLandscape
                    OccFPrincipale.ActiveForm.PrintForm
                    Printer.Orientation = vbPRORPortrait
                
                Case Else
            End Select
        
        Case FENETRES.F_TRACABILITE_PRODUCTION
            '--- op�rations � effectuer ---
            Select Case EtatSouhaite
                Case TYPES_IMPRESSIONS.TI_APERCU_AVANT_IMPRESSION, TYPES_IMPRESSIONS.TI_IMPRIMER

                    '--- contr�les et appel de l'�cran des impressions ---
                    'If PossibiliteImpression = False Then Exit Sub
                    OptionImpressionChoisie = ChoixImpression(NumfenetreActive, _
                                                                                              "Par le n� de fiche de traitement", _
                                                                                              "La fiche de traitement par la commande interne", _
                                                                                              "La production par la date d'entr�e en ligne")

                    '--- affectation des crit�res d'impression ---
                    Select Case OptionImpressionChoisie
                        
                        Case 1
                            '--- par le n� de traitement (fiche de production) ---
                            TCriteresImpression(1) = OccFPrincipale.ActiveForm.ADODCDetailsChargesProduction.Recordset("NumFicheProduction")
                            If TCriteresImpression(1) <> "" Then
                                ConstructionImpressionDetailsCharge NumFicheProduction:=TCriteresImpression(1), _
                                                                                              NumCommandeInterne:=""
                                ConstructionImpressionGammesProduction NumFicheProduction:=TCriteresImpression(1)
                                ConstructionImpressionTracabiliteCharge NumFicheProduction:=TCriteresImpression(1)
                                ConstructionImpressionAlarmesLigne NumFicheProduction:=TCriteresImpression(1)
                            End If
                        
                        Case 2
                            '--- par le n� de commande interne ---
                            With OccFPrincipale.ActiveForm.ADODCDetailsChargesProduction
                                TCriteresImpression(1) = .Recordset("NumFicheProduction")
                                TCriteresImpression(2) = .Recordset("NumCommandeInterne")
                            End With
                            If TCriteresImpression(1) <> "" And TCriteresImpression(2) <> "" Then
                                ConstructionImpressionDetailsCharge NumFicheProduction:=TCriteresImpression(1), _
                                                                                                NumCommandeInterne:=TCriteresImpression(2)
                                ConstructionImpressionGammesProduction NumFicheProduction:=TCriteresImpression(1)
                                ConstructionImpressionTracabiliteCharge NumFicheProduction:=TCriteresImpression(1)
                                ConstructionImpressionAlarmesLigne NumFicheProduction:=TCriteresImpression(1)
                            End If
                        
                        Case 3
                            '--- la production par la date d'entr�e en ligne ---
                            With OccFPrincipale.ActiveForm
                                TCriteresImpression(1) = .TBCommencantPar.Text
                            End With
                       
                        Case Else
                    End Select
                    
                    '--- lancement de l'aper�u ou de l'impression ---
                    Call OccFPrincipale.ActiveForm.SourisEnAttente(True)
                    If EtatSouhaite = TYPES_IMPRESSIONS.TI_APERCU_AVANT_IMPRESSION Then
                        
                        Select Case OptionImpressionChoisie
                            
                            Case 1, 2
                                '--- par le n� de traitement (fiche de production) et le n� de commande interne ---
                                DRTra�abilite_DetailsCharge1.Show vbModal, OccFPrincipale
                                DRTra�abilite_GammesProduction1.Show vbModal, OccFPrincipale
                                DRTra�abilite_Tra�abiliteCharge1.Show vbModal, OccFPrincipale
                                DRTra�abilite_AlarmesLigne1.Show vbModal, OccFPrincipale
                            
                            Case 3
                                '--- la production par la date d'entr�e en ligne ---
                                If TCriteresImpression(1) <> "" Then
                                    If IsDate(TCriteresImpression(1)) = True Then
                                        DRTra�abilite_ProductionParJour1.Show vbModal, OccFPrincipale
                                    End If
                                End If
                            
                            Case Else
                        End Select
                    
                    Else
                        
                        Select Case OptionImpressionChoisie
                            
                            Case 1, 2
                                '--- par le n� de traitement (fiche de production) et le n� de commande interne ---
                                DRTra�abilite_DetailsCharge1.PrintReport
                                Pause 5
                                DRTra�abilite_GammesProduction1.PrintReport
                                Pause 5
                                DRTra�abilite_Tra�abiliteCharge1.PrintReport
                                Pause 5
                                DRTra�abilite_AlarmesLigne1.PrintReport
                            
                            Case 3
                                '--- la production par la date d'entr�e en ligne ---
                                If TCriteresImpression(1) <> "" Then
                                    If IsDate(TCriteresImpression(1)) = True Then
                                        DRTra�abilite_ProductionParJour1.PrintReport
                                    End If
                                End If
                            
                            Case Else
                        End Select
                    
                    End If
                    Call OccFPrincipale.ActiveForm.SourisEnAttente(False)

                Case TYPES_IMPRESSIONS.TI_IMPRIMER_FENETRE_ACTIVE
                    '--- impression de la fen�tre active ---
                    Printer.Orientation = vbPRORLandscape
                    OccFPrincipale.ActiveForm.PrintForm
                    Printer.Orientation = vbPRORPortrait

                Case Else
            End Select

        Case Else

    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Choix de l'impression
' Entr�es :  NumfenetreAppel -> Num�ro de fenetre ayant lanc� l'appel
'                   TypeImpression -> Type d'impression
'                        TexteOption1 -> Texte de l'option 1
'                        TexteOption2 -> Texte de l'option 2
'                        TexteOption3 -> Texte de l'option 3
'                        TexteOption4 -> Texte de l'option 4
'                        TexteOption5 -> Texte de l'option 5
'                        TexteOption6 -> Texte de l'option 6
'                        TexteOption7 -> Texte de l'option 7
'                        TexteOption8 -> Texte de l'option 8
' Retours :  ChoixImpression -> 0 = annuler
'                                                  Autres valeurs = num�ro de l'option s�lectionn�
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ChoixImpression(ByVal NumfenetreAppel As Long, _
                                                      Optional ByVal TexteOption1 As Variant, _
                                                      Optional ByVal TexteOption2 As Variant, _
                                                      Optional ByVal TexteOption3 As Variant, _
                                                      Optional ByVal TexteOption4 As Variant, _
                                                      Optional ByVal TexteOption5 As Variant, _
                                                      Optional ByVal TexteOption6 As Variant, _
                                                      Optional ByVal TexteOption7 As Variant, _
                                                      Optional ByVal TexteOption8 As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    'Static PremierChargement As Boolean
    Dim a As Integer
    Dim TTextesOptions(1 To 8) As String
    
    '--- remplissage du tableau ---
    If IsMissing(TexteOption1) = False Then TTextesOptions(1) = TexteOption1
    If IsMissing(TexteOption2) = False Then TTextesOptions(2) = TexteOption2
    If IsMissing(TexteOption3) = False Then TTextesOptions(3) = TexteOption3
    If IsMissing(TexteOption4) = False Then TTextesOptions(4) = TexteOption4
    If IsMissing(TexteOption5) = False Then TTextesOptions(5) = TexteOption5
    If IsMissing(TexteOption6) = False Then TTextesOptions(6) = TexteOption6
    If IsMissing(TexteOption7) = False Then TTextesOptions(7) = TexteOption7
    If IsMissing(TexteOption8) = False Then TTextesOptions(8) = TexteOption8
    
    '--- chargement de la fenetre ---
    'If PremierChargement = False Then
        Load FChoixImpression
        'PremierChargement = True
    'End If
    
    With FChoixImpression
        
        '--- affectation ---
        .NumFenetre = FENETRES.F_CHOIX_IMPRESSION
        .NumfenetreAppel = NumfenetreAppel
       
        '--- textes des options ---
        For a = LBound(TTextesOptions()) To UBound(TTextesOptions())

            If TTextesOptions(a) <> "" Then
                
                With .OBOptionsImpression(a)
                    .Caption = TTextesOptions(a)
                    .Enabled = True
                End With
            
                With .TBMargesGauche(a)
                    .BackColor = COULEURS.BLANC
                    .Text = "0,00"
                    .Enabled = True
                End With
                
                With .TBMargesHaute(a)
                    .BackColor = COULEURS.BLANC
                    .Text = "0,00"
                    .Enabled = True
                End With
                
                With .TBMargesDroite(a)
                    .BackColor = COULEURS.BLANC
                    .Text = "0,00"
                    .Enabled = True
                End With
                
                With .TBMargesBasse(a)
                    .BackColor = COULEURS.BLANC
                    .Text = "0,00"
                    .Enabled = True
                End With
            
            Else
                
                With .OBOptionsImpression(a)
                    .Caption = ""
                    .Enabled = False
                End With
            
                With .TBMargesGauche(a)
                    .BackColor = COULEURS.GRIS_2
                    .Text = ""
                    .Enabled = False
                End With
                
                With .TBMargesHaute(a)
                    .BackColor = COULEURS.GRIS_2
                    .Text = ""
                    .Enabled = False
                End With
                
                With .TBMargesDroite(a)
                    .BackColor = COULEURS.GRIS_2
                    .Text = ""
                    .Enabled = False
                End With
                
                With .TBMargesBasse(a)
                    .BackColor = COULEURS.GRIS_2
                    .Text = ""
                    .Enabled = False
                End With
            
            End If
        
        Next a
    
        '--- affichage de la fenetre ---
        .Show vbModal, OccFPrincipale
    
        '--- analyse de la r�ponse ---
        ChoixImpression = .OptionSelectionnee
    
    End With
    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Centre un �tat sur une fenetre pour l'impression
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub CentreEtat(ByRef EtatAImprimer As DataReport)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- centrage de la page ---
    With EtatAImprimer
        'Printer.ScaleLeft = -((Printer.Width - .ReportWidth) / 2)
        'Printer.ScaleTop = -((Printer.Height - .Height) / 2)
        .LeftMargin = ((Printer.Width - .ReportWidth) / 2)
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Contr�le la possibilit� d'impression
' Entr�es :
' Retours : PossibiliteImpression -> FALSE = Impression non possible
'                                                         TRUE = Impression possible
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function PossibiliteImpression() As Boolean
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    Const TITRE_MESSAGE_IMPRESSION As String = " Demande d'impression"

    '--- affectation ---
    PossibiliteImpression = True
    
    '--- contr�le ---
    If OccFPrincipale.ActiveForm.CBValider.Enabled = True Then
        MessageErreur TITRE_MESSAGE_IMPRESSION, MESSAGE_210
        PossibiliteImpression = False
    End If

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affecte les marges sur un �tat � imprimer
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub MargesEtat(ByRef EtatAImprimer As DataReport)
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- centrage de la page ---
    With EtatAImprimer
        .LeftMargin = MargeGaucheTwips
        .TopMargin = MargeHauteTwips
        .RightMargin = MargeDroiteTwips
        .BottomMargin = MargeBasseTwips
    End With
                    
End Sub



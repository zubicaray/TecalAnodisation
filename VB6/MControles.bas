Attribute VB_Name = "MControles"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES CONTROLES
' Nom                    : MControles.bas
' Date de création : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- pour les contrôles des températures ---
Public Enum CONTROLES_TEMPERATURES
    C_PAS_DE_CONTROLE = 0                  'pas de contrôles sur les températures (cas de l'arrêt et de la veille)
    C_TEMPERATURE_NORMALE = 1       'température considérée comme normale
    C_TEMPERATURE_INFERIEURE = 2   'la température mesurée est inférieure au seuil minimum
    C_TEMPERATURE_SUPERIEURE = 3  'la température mesurée est supérieure au seuil maximum
    C_DEFAUT_PT100 = 4                          'il y a un défaut sur la PT 100
    C_TEMPERATURE_CRITIQUE = 5        'la température critique vient d'être dépassée
End Enum

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Contrôle la nullité d'un champ d'une table ou requête d'une base de donnée
' Entrées :     Enregistrement -> Enregistrement d'une table
'                          NomChamp -> Nom du champ à contrôler
'                     ValeurSiNullite -> Valeur à affecter si le champ est nul
' Retours : C_Nullite_Champ -> Valeur du champ si non nul sinon ValeurSiNullite
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function C_Nullite_Champ(ByVal Enregistrement As ADODB.Recordset, _
                                                       ByVal NomChamp As String, _
                                                       ByVal ValeurSiNullite As Variant) As Variant
      
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim ValeurChamp As Variant
    
    '--- affectation ---
    ValeurChamp = Enregistrement.Fields(NomChamp).value
    
    '--- contrôle ---
    If IsNull(ValeurChamp) = True Or ValeurChamp = Null Then
        C_Nullite_Champ = ValeurSiNullite
    Else
        C_Nullite_Champ = ValeurChamp
    End If
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Controle la température d'une cuve
' Entrées :                   NumCuve -> Numéro de la cuve concerné
' Retours : ControleTemperature -> valeurs de l'énumération CONTROLES_TEMPERATURES
'                                                        0 = pas de controles sur les températures (cas de l'arrêt et de la veille)
'                                                        1 = température considérée comme normale
'                                                        2 = la température mesurée est inférieure au seuil minimum
'                                                        3 = la température mesurée est supérieure au seuil maximum
'                                                        4 = il y a un défaut sur la PT 100
'                                                        5 = la température critique vient d'être dépassée
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControleTemperature(ByVal NumCuve As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    Dim TempComparaison As Single
    Static TTempAtteinteUneFois(CUVES_API.C_C00 To DERNIERE_CUV_REGULATION) As Boolean

    '--- affectation ---
    ControleTemperature = CONTROLES_TEMPERATURES.C_PAS_DE_CONTROLE

    '--- contrôle ---
    With TEtatsCuvesAPI(NumCuve).Temperatures

        If .TempActuelle = 32767 / 10 Or .TempActuelle = -32768 / 10 Then

            '--- cas du défaut de la PT 100 ---
            TEtatsCuvesAPI(NumCuve).TEntreesAPI.DefautPT100 = True
            ControleTemperature = CONTROLES_TEMPERATURES.C_DEFAUT_PT100
            Exit Function

        Else

            '--- recherche de la température en fonction du mode de production en cours ---
            Select Case TEtatsCuvesAPI(NumCuve).ModeProduction

                Case MODES_PRODUCTION.M_ARRET
                    '--- mode arrêt ---
                    TTempAtteinteUneFois(NumCuve) = False
                    
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropBasse = False
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropHaute = False
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.DefautPT100 = False
                    
                    Exit Function

                Case MODES_PRODUCTION.M_VEILLE
                    '--- mode veille ---
                    TTempAtteinteUneFois(NumCuve) = False
                    
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropBasse = False
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropHaute = False
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.DefautPT100 = False
                    
                    Exit Function

                Case MODES_PRODUCTION.M_PRODUCTION
                    '--- mode production ---
                    TempComparaison = .TempProduction
                    If .TempActuelle >= .TempProduction Then
                        TTempAtteinteUneFois(NumCuve) = True
                    End If

                Case Else: Exit Function
            
            End Select

            '--- affectation ---
            ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_NORMALE  'température normale par défaut

            '--- comparaisons ---
            If TTempAtteinteUneFois(NumCuve) = True Then
                If .TempActuelle < (TempComparaison - .EcartInferieurAlarme) Then
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropBasse = True
                    ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_INFERIEURE
                End If
                If .TempActuelle > (TempComparaison + .EcartSuperieurAlarme) Then
                    TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropHaute = True
                    ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_SUPERIEURE
                End If
            End If

        End If

    End With

    '--- pas de défaut si la température est normale ---
    If ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_NORMALE Then
        TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropBasse = False
        TEtatsCuvesAPI(NumCuve).TEntreesAPI.TemperatureTropHaute = False
        TEtatsCuvesAPI(NumCuve).TEntreesAPI.DefautPT100 = False
    End If

End Function


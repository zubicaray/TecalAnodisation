Attribute VB_Name = "MControles"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DES CONTROLES
' Nom                    : MControles.bas
' Date de cr�ation : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- pour les contr�les des temp�ratures ---
Public Enum CONTROLES_TEMPERATURES
    C_PAS_DE_CONTROLE = 0                  'pas de contr�les sur les temp�ratures (cas de l'arr�t et de la veille)
    C_TEMPERATURE_NORMALE = 1       'temp�rature consid�r�e comme normale
    C_TEMPERATURE_INFERIEURE = 2   'la temp�rature mesur�e est inf�rieure au seuil minimum
    C_TEMPERATURE_SUPERIEURE = 3  'la temp�rature mesur�e est sup�rieure au seuil maximum
    C_DEFAUT_PT100 = 4                          'il y a un d�faut sur la PT 100
    C_TEMPERATURE_CRITIQUE = 5        'la temp�rature critique vient d'�tre d�pass�e
End Enum

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Contr�le la nullit� d'un champ d'une table ou requ�te d'une base de donn�e
' Entr�es :     Enregistrement -> Enregistrement d'une table
'                          NomChamp -> Nom du champ � contr�ler
'                     ValeurSiNullite -> Valeur � affecter si le champ est nul
' Retours : C_Nullite_Champ -> Valeur du champ si non nul sinon ValeurSiNullite
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function C_Nullite_Champ(ByVal Enregistrement As ADODB.Recordset, _
                                                       ByVal NomChamp As String, _
                                                       ByVal ValeurSiNullite As Variant) As Variant
      
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim ValeurChamp As Variant
    
    '--- affectation ---
    ValeurChamp = Enregistrement.Fields(NomChamp).Value
    
    '--- contr�le ---
    If IsNull(ValeurChamp) = True Or ValeurChamp = Null Then
        C_Nullite_Champ = ValeurSiNullite
    Else
        C_Nullite_Champ = ValeurChamp
    End If
    
End Function
' SZ 202110
Public Function getCuveId_OLD(ByVal IdxAutomate As Integer) As Integer
    'Dim I As Long
    'Call Log("------------------------------")
    'Call Log("idx auto:" & IdxAutomate)
    

    
    
     'Call Log("idx cuve: -1 ")
    
    getCuveId_OLD = CORRESPONDANCES_IDX_CUVES_API(IdxAutomate)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Controle la temp�rature d'une cuve
' Entr�es :                   NumCuve -> Num�ro de la cuve concern�
' Retours : ControleTemperature -> valeurs de l'�num�ration CONTROLES_TEMPERATURES
'                                                        0 = pas de controles sur les temp�ratures (cas de l'arr�t et de la veille)
'                                                        1 = temp�rature consid�r�e comme normale
'                                                        2 = la temp�rature mesur�e est inf�rieure au seuil minimum
'                                                        3 = la temp�rature mesur�e est sup�rieure au seuil maximum
'                                                        4 = il y a un d�faut sur la PT 100
'                                                        5 = la temp�rature critique vient d'�tre d�pass�e
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ControleTemperature(ByVal NumCuve As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    Dim TempComparaison As Single
    Static TTempAtteinteUneFois(CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As Boolean

    '--- affectation ---
    ControleTemperature = CONTROLES_TEMPERATURES.C_PAS_DE_CONTROLE

    '--- contr�le ---
    With TEtatsCuves(NumCuve).Temperatures

        If .TempActuelle = 32767 / 10 Or .TempActuelle = -32768 / 10 Then

            '--- cas du d�faut de la PT 100 ---
            TEtatsCuves(NumCuve).TEntreesAPI.DefautPT100 = True
            ControleTemperature = CONTROLES_TEMPERATURES.C_DEFAUT_PT100
            Exit Function

        Else

            '--- recherche de la temp�rature en fonction du mode de production en cours ---
            Select Case TEtatsCuves(NumCuve).ModeProduction

                Case MODES_PRODUCTION.M_ARRET
                    '--- mode arr�t ---
                    TTempAtteinteUneFois(NumCuve) = False
                    
                    TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropBasse = False
                    TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropHaute = False
                    TEtatsCuves(NumCuve).TEntreesAPI.DefautPT100 = False
                    
                    Exit Function

                Case MODES_PRODUCTION.M_VEILLE
                    '--- mode veille ---
                    TTempAtteinteUneFois(NumCuve) = False
                    
                    TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropBasse = False
                    TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropHaute = False
                    TEtatsCuves(NumCuve).TEntreesAPI.DefautPT100 = False
                    
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
            ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_NORMALE  'temp�rature normale par d�faut

            '--- comparaisons ---
            If TTempAtteinteUneFois(NumCuve) = True Then
                If .TempActuelle < (TempComparaison - .EcartInferieurAlarme) Then
                    TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropBasse = True
                    ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_INFERIEURE
                End If
                If .TempActuelle > (TempComparaison + .EcartSuperieurAlarme) Then
                    TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropHaute = True
                    ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_SUPERIEURE
                End If
            End If

        End If

    End With

    '--- pas de d�faut si la temp�rature est normale ---
    If ControleTemperature = CONTROLES_TEMPERATURES.C_TEMPERATURE_NORMALE Then
        TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropBasse = False
        TEtatsCuves(NumCuve).TEntreesAPI.TemperatureTropHaute = False
        TEtatsCuves(NumCuve).TEntreesAPI.DefautPT100 = False
    End If

End Function


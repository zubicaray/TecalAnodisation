Attribute VB_Name = "MPFenetres"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE CONTENANT LES OCCURRENCES DES FENETRES
' Nom                    : MPFenetres.bas
' Date de cr�ation : 31/07/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- occurrence des fenetres du programmes ---
Public OccFPrincipale As New FPrincipale                                                                                            'fen�tre principale
Public OccFSynoptique As New FSynoptique                                                                                        'fen�tre du synoptique

Public OccFOrganisationLigne As New FOrganisationLigne                                                                 'fen�tre de l'organisation de la ligne
Public OccFMoteurInference As New FMoteurInference                                                                        'fen�tre du moteur d'inf�rence
Public OccFPremisses As New FPremisses                                                                                          'fen�tre des pr�misses
Public OccFTempsMouvements As New FTempsMouvements                                                             'fen�tre des temps de mouvements
Public OccFModeCyclique As New FModeCyclique                                                                               'fen�tre du mode cyclique

Public OccFGammesAnodisation As New FGammesAnodisation                                                          'gammes d'anodisation
Public OccFTra�abiliteProduction As New FTra�abiliteProduction                                                        'tra�abilit� de la production
Public OccFVisualisationGraphesProduction As New FVisualisationGraphesProduction                     'visualisation des graphes de production
Public OccFNettoyageGraphesProduction As New FNettoyageGraphesProduction                              'nettoyage des graphes de production

Public OccFChargesEnLigne As New FChargesEnLigne                                                                       'fen�tre des charges en ligne

Public OccFCyclesPonts As New FCyclesPonts                                                                                    'fen�tre des cycles des ponts
Public OccFChargementPrevisionnel As New FChargementPrevisionnel                                             'fen�tre du chargement et du pr�visonnel

Public OccFGestionRedresseurs As New FGestionRedresseurs                                                          'fen�tre de la gestion des redresseurs

Public OccFGestionCuves As New FGestionCuves                                                                               'fen�tre de la gestion des cuves
Public OccFGestionRegulation As New FGestionRegulation                                                                 'fen�tre de la gestion de la r�gulation
Public OccFProgrammateurCyclique As New FProgrammateurCyclique                                               'fen�tre du programmateur cyclique
Public OccFAnnexes As New FAnnexes                                                                                                 'fen�tre des annexes
Public OccFListeDefauts As New FListeDefauts                                                                                    'fen�tre contenant la liste des d�fauts

Public OccFMaintenance As New FMaintenance                                                                                    'fen�tre de la maintenance
Public OccFInformationsDefautsVariateurs As New FInformationsDefautsVariateurs                           'fen�tre des informations sur les d�fauts des variateurs
Public OccFInformationsDefautsCommunicationAutomate As New FInformationsDefautsCommunicationAutomate       'fen�tre des informations sur les d�fauts de communication avec l'automate

Public OccFTra�abiliteAlarmes As New FTra�abiliteAlarmes                                                                 'tra�abilit� des alarmes
Public OccFAPropos As New FAPropos                                                                                                  'fen�tre � propos

Public OccFEssais As New FEssais                                                                                                      'fen�tre pour les essais




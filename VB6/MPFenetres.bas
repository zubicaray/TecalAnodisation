Attribute VB_Name = "MPFenetres"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE CONTENANT LES OCCURRENCES DES FENETRES
' Nom                    : MPFenetres.bas
' Date de création : 31/07/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- occurrence des fenetres du programmes ---
Public OccFPrincipale As New FPrincipale                                                                                            'fenêtre principale
Public OccFSynoptique As New FSynoptique                                                                                        'fenêtre du synoptique

Public OccFOrganisationLigne As New FOrganisationLigne                                                                 'fenêtre de l'organisation de la ligne
Public OccFMoteurInference As New FMoteurInference                                                                        'fenêtre du moteur d'inférence
Public OccFPremisses As New FPremisses                                                                                          'fenêtre des prémisses
Public OccFTempsMouvements As New FTempsMouvements                                                             'fenêtre des temps de mouvements
Public OccFModeCyclique As New FModeCyclique                                                                               'fenêtre du mode cyclique

Public OccFGammesAnodisation As New FGammesAnodisation                                                          'gammes d'anodisation
Public OccFTraçabiliteProduction As New FTraçabiliteProduction                                                        'traçabilité de la production
Public OccFVisualisationGraphesProduction As New FVisualisationGraphesProduction                     'visualisation des graphes de production
Public OccFNettoyageGraphesProduction As New FNettoyageGraphesProduction                              'nettoyage des graphes de production

Public OccFChargesEnLigne As New FChargesEnLigne                                                                       'fenêtre des charges en ligne

Public OccFCyclesPonts As New FCyclesPonts                                                                                    'fenêtre des cycles des ponts
Public OccFChargementPrevisionnel As New FChargementPrevisionnel                                             'fenêtre du chargement et du prévisonnel

Public OccFGestionRedresseurs As New FGestionRedresseurs                                                          'fenêtre de la gestion des redresseurs

Public OccFGestionCuves As New FGestionCuves                                                                               'fenêtre de la gestion des cuves
Public OccFGestionRegulation As New FGestionRegulation                                                                 'fenêtre de la gestion de la régulation
Public OccFProgrammateurCyclique As New FProgrammateurCyclique                                               'fenêtre du programmateur cyclique
Public OccFAnnexes As New FAnnexes                                                                                                 'fenêtre des annexes
Public OccFListeDefauts As New FListeDefauts                                                                                    'fenêtre contenant la liste des défauts

Public OccFMaintenance As New FMaintenance                                                                                    'fenêtre de la maintenance
Public OccFInformationsDefautsVariateurs As New FInformationsDefautsVariateurs                           'fenêtre des informations sur les défauts des variateurs
Public OccFInformationsDefautsCommunicationAutomate As New FInformationsDefautsCommunicationAutomate       'fenêtre des informations sur les défauts de communication avec l'automate

Public OccFTraçabiliteAlarmes As New FTraçabiliteAlarmes                                                                 'traçabilité des alarmes
Public OccFAPropos As New FAPropos                                                                                                  'fenêtre à propos

Public OccFEssais As New FEssais                                                                                                      'fenêtre pour les essais




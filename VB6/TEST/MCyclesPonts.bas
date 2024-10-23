Attribute VB_Name = "MCyclesPonts"
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE GERANT LE CYCLE DU PONT (ACTIONS)
' Nom                    : MCyclePont.bas
' Date de création : 20/11/2006
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

' --- déclarations obligatoires ---
Option Explicit

' --- options générales ---
Option Base 1
DefVar A-Z

'--- tableaux publiques ---

'--- cycle du pont ---
Public TCyclesPonts(PONTS.P_1 To PONTS.P_2) As VarCyclePont

' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Retourne une action complète du PONT à partir de son numéro
' Entrées : NumAction -> n° de l'action concerné
' Retours :
' Détails  :
' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ActionsPont(ByVal NumAction As Long) As EnrActions

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    Const NOP As Integer = 0                                                                                'pas d'opération
    
    Const DEBUT_POS_POSTES As Integer = 0  'POSTES.P_POLISSEUSE                     'début des positions des postes
    Const FIN_POS_POSTES As Integer = 10 'POSTES.P_DECHARGEMENT                     'fin des positions des postes
    
    Const RECALAGE_PONT As Integer = 50                                         'recalage du pont sur les capteurs d'initialisation
        
    Const RELACHEMENT_FREINS_TRL As Integer = 70                                'relachement des freins des 2 moteurs de la translation

    Const TRL_POSTE_SELECTIONNE As Integer = 100                                'translation à un poste sélectionné
    
    Const PERMUT_CUVE_POSTE_CR As Integer = 110                                 'permutation de la cuve et du poste de chrome
        
    Const POSTE_OU_CHARGER As Integer = 120                                     'translation au poste de chargement choisi (chargeur automatique ou polisseuse)
        
    Const DEBUT_TRL_PERMUT As Integer = 150                                     'début de la translation à la permutation
    Const FIN_TRL_PERMUT As Integer = 151                                       'fin de la translation à la permutation

    Const NIVEAU_BAS As Integer = 201                                           'code du NIVEAU BAS
    Const NIVEAU_HAUT As Integer = 215                                          'code du NIVEAU HAUT

    Const FORCER_MONTEE_EN_HAUT As Integer = 260                                'force la montée pour atteindre le capteur haut

    Const FORCER_DESCENTE_INTER As Integer = 270                                'force la descente pour atteindre le capteur intermédiaire
    
    Const FORCER_REF_LEVAGE As Integer = 280                                    'force la descente pour atteindre le capteur bas et effectue le référence codeur

    Const DEBUT_TEMPO As Integer = 300                                          'début de l'action de temporisation
    Const FIN_TEMPO As Integer = 399                                            'fin de l'action de temporisation
    
    Const TEMPO_EGOUTTAGE As Integer = 410                                      'temporisation d'égouttage

    Const TEMPO_DEGRAISSAGE As Integer = 450                                    'temporisation de dégraissage

    Const TEMPO_RINCAGE As Integer = 460                                        'temporisation au rinçage

    Const MONTEE_IMPULSION_CHARG As Integer = 500                               'demande de MONTEE à un niveau conditionné par une impulsion
                                                                                'pour le CHARGEMENT

    Const DESCENTE_IMPULS_DECHARG As Integer = 510                              'demande de DESCENTE à un niveau conditionné par une impulsion
                                                                                'pour le DECHARGEMENT

    Const ATTENTE_AUTOR_DEPL_DECH As Integer = 520                              'attente de l'autorisation de DEPLACEMENT au POSTE de déchargement

    Const SORTIE_BACS_ANTI_EGOUT As Integer = 600                               'sortie des bacs anti-égouttures
    Const GARAGE_BACS_ANTI_EGOUT As Integer = 610                               'position garage des bacs anti-égouttures
    Const CTRL_BACS_ANTI_EGOUT As Integer = 620                                 'contrôle de la position garage des bacs anti-égouttures
        
    Const SYNCHRO_CHARGEMENT_AUTO As Integer = 900                              'SYNCHRO avec la PHASE de CHARGEMENT 1er descente (Polisseuse, Listo, Automatique)
    Const SYNCHRO_CHARGEUR_AUTO As Integer = 910                                'SYNCHRO avec le CHARGEUR AUTOMATIQUE MODE NORMAL, demande de mise en position du chargeur
    
    Const APPEL_NIVEAUX_CHARGEMENT As Integer = 1040                            'appel de la table des niveaux du POSTE de CHARGEMENT
    Const APPEL_NIVEAUX_DECHARG As Integer = 1050                               'appel de la table des niveaux du POSTE de DECHARGEMENT
    
    Const RAZ_CHARGE As Integer = 1070                                          'RAZ de la CHARGE

    Const ACQUIS_OF As Integer = 1080                                           'acquisition de l'ordre de fabrication

    Const AFFECT_CHARGE_PONT As Integer = 1090                                  'affectation du numéro de charge sur le pont au chargement

    Const ATTENTE_VALID_APRES_CHAR As Integer = 2001                            'attente de la validation après le chargement
    Const ATTENTE_DESCENTE_DEGRAIS As Integer = 2002                            'attente pour la descente au poste de dégraissage
    Const ATTENTE_AU_DESSUS_CHROME As Integer = 2003                            'attente au dessus du chromage si le pont est arrivée trop tôt
    Const ATTENTE_DESCENTE_ATTAQUE As Integer = 2004                            'attente pour la descente au poste d'attaque

    Const CTRL_COMMUT_PASS_DEGRAIS As Integer = 2005                            'contrôle du commutateur de passage au dégraissage
    Const CTRL_COMMUT_PASS_ATTAQUE As Integer = 2006                            'contrôle du commutateur de passage à l'attaque

    Const ATTENTE_FIN_CHROMAGE As Integer = 2010                                'attente fin du cycle de chromage pour descente au poste

    Const ATTENTE_ARRET_REDRESSEUR As Integer = 2011                            'attente de l'arrêt d'un redresseur (contrôle sur l'intensité)

    Const CRTL_SUIVI_AV_DESC As Integer = 2500                                  'contrôle du suivi avant descente à un poste
    
    Const CTRL_DEM_CHANGE_OUTIL As Integer = 3000                               'contrôle de la demande de changement outil
    Const DEM_SORTIE_TRANSFERT As Integer = 3001                                'demande de SORTIE du chariot de transfert (EN LIGNE)
    Const DEM_RENTREE_TRANSFERT As Integer = 3002                               'demande de RENTREE du chariot de transfert (HORS LIGNE)
    Const DEM_BON_NIVEAU_TRANFERT As Integer = 3003                             'demande de mise au bon niveau du chariot de transfert

    Const CTRL_TRANSFERT_SORTIE As Integer = 3011                               'contrôle chariot de transfert SORTIE (EN LIGNE)
    Const CTRL_TRANSFERT_RENTRE As Integer = 3012                               'contrôle chariot de transfert RENTRE (HORS LIGNE)
    Const CTRL_BON_NIV_TRANSFERT As Integer = 3013                              'contrôle du bon niveau du chariot de transfert

    Const PRISE_DEPOSE_TRANSFERT As Integer = 3050                              'affectation des postes de dépose et de prise
    Const TRL_POSTE_DEPOSE_TRANS As Integer = 3060                              'translation au poste de DEPOSE sur le chariot de transfert (emplacement vide)
    Const TRL_POSTE_PRISE_TRANS As Integer = 3070                               'translation au poste de PRISE sur le chariot de transfert
        
    Const ATTENTE_VALID_OUTIL As Integer = 3100                                 'attente du code validation du nouvel outil

    Const CTRL_FORCER_MANUEL As Integer = 3110                                  'contrôle de forçage en manuel du pont

    Const CTRL_DESCENTE_DEGRAIS As Integer = 4000                               'contrôle si la descente est possible (poste 1) du dégraissage
    Const DEM_DEMI_POSTE_DEGRAIS As Integer = 4010                              'demande d'avance au demi poste du dégraissage
    Const LANCEMENT_TEMPO_DEGRAIS As Integer = 4020                             'lancement de la temporisation de dégraissage
    Const ARRET_TRL_DEGRAISSAGE As Integer = 4030                               'arrêt des 2 translateurs du dégraissage
    Const REF_AXE_TRL_DEGRAISSAGE As Integer = 4040                             'référence d'axes des 2 translateurs du dégraissage
    
    Const CTRL_DESCENTE_ATTAQ As Integer = 4400                                 'contrôle si la descente est possible à l'attaque
    Const ATTENTE_FIN_ATTAQUE As Integer = 4500                                 'attente de la fin de l'attaque
    Const AUTOR_ATTAQUE_MODE_2P As Integer = 4550                               'autorisation de lancement de l'attaque en mode 2 ponts

    Const CTRL_DESCENTE_REPRISE As Integer = 4600                               'contrôle si la descente est possible au poste de reprise

    Const CHOIX_CYCLE As Integer = 5000                                         'choix du cycle (chromage ou dégraissage)
    Const CHOIX_MODE_CHARGEMENT As Integer = 5200                               'Choix du mode de chargement avec INIT (Polisseuse, Forcer Listo OU Chargeur), code
                                                                                'Aiguillage +1 Mode chargeur automatique, Aiguillage +2 Autres mode

    Const LANCE_DECH_AUTO As Integer = 6000                                     'lancement du déchargement automatique
        
    Const DEBUT_SYNCHRO_P1 As Integer = 7001                                    'début de la synchronisation relatif au pont 1 (synchros de 1 à 10)
    Const FIN_SYNCHRO_P1 As Integer = 7010                                      'fin de la synchronisation relatif au pont 1 (synchros de 1 à 10)

    Const DEBUT_SYNCHRO_P2 As Integer = 7501                                    'début de la synchronisation relatif au pont 2 (synchros de 1 à 10)
    Const FIN_SYNCHRO_P2 As Integer = 7510                                      'fin de la synchronisation relatif au pont 2 (synchros de 1 à 10)
    
    Const FCY As Integer = 8000                                                 'code de fin de cycle

    Const DEBUT_SEQ_SAUT As Integer = 10000                                     'début de la séquence de saut (équivalent GOTO)
    Const FIN_SEQ_SAUT As Integer = 10299                                       'fin de la séquence de saut (équivalent GOTO)

    '--- déclaration ---
    
    '--- analyse en fonction du numéro de l'action ---
    With ActionsPont
        
        Select Case NumAction
    
            Case NOP
                '--- pas d'opération ---
                .CodeAction = "NOP"
                .LibelleAction = "Pas d'opération"
            
            Case DEBUT_POS_POSTES To FIN_POS_POSTES
                '--- translation directe au poste ---
                .CodeAction = "TRANSLATION DIRECTE"
                .LibelleAction = "Translation au poste " & NumAction & " - " & TEtatsPostes(NumAction).DefinitionPoste.LibellePoste
            
            Case RECALAGE_PONT
                '--- recalage du pont sur les capteurs d'initialisation ---
                .CodeAction = "RECALAGE_PONT"
                .LibelleAction = "Recalage du pont sur les capteurs d'initialisation"
        
            Case RELACHEMENT_FREINS_TRL
                '--- relachement des freins des 2 moteurs de la translation ---
                .CodeAction = "RELACHEMENT_FREINS_TRL"
                .LibelleAction = "Relachement des freins des 2 moteurs de la translation"
            
            Case TRL_POSTE_SELECTIONNE
                '--- translation à un poste sélectionné ---
                .CodeAction = "TRL_POSTE_SELECTIONNE"
                .LibelleAction = "Translation à un poste sélectionné"
        
            Case PERMUT_CUVE_POSTE_CR
                '--- permutation de la cuve et du poste de chrome ---
                .CodeAction = "PERMUT_CUVE_POSTE_CR"
                .LibelleAction = "Permutation de la cuve et du poste de chrome"
            
            Case POSTE_OU_CHARGER
                '--- translation au poste de chargement choisi (chargeur automatique ou polisseuse) ---
                .CodeAction = "POSTE_OU_CHARGER"
                .LibelleAction = "Translation au poste de chargement choisi (chargeur automatique ou polisseuse)"
        
            Case DEBUT_TRL_PERMUT To FIN_TRL_PERMUT
                '--- translation à la permutation ---
                .CodeAction = "TRL_PERMUT"
                .LibelleAction = "Translation à la permutation"
            
            Case NIVEAU_BAS To NIVEAU_HAUT
                '--- atteindre un niveau ---
                .CodeAction = "NIVEAU"
                .LibelleAction = "Atteindre le niveau " & NumAction - NIVEAU_BAS + 1
    
            Case FORCER_MONTEE_EN_HAUT
                '--- force la montée pour atteindre le capteur haut ---
                .CodeAction = "FORCER_MONTEE_EN_HAUT"
                .LibelleAction = "Force la montée pour atteindre le capteur haut"
            
            Case FORCER_DESCENTE_INTER
                '--- force la descente pour atteindre le capteur intermédiaire ---
                .CodeAction = "FORCER_DESCENTE_INTER"
                .LibelleAction = "Force la descente pour atteindre le capteur intermédiaire"
            
            Case FORCER_REF_LEVAGE
                '--- force la descente pour atteindre le capteur de niveau bas et effectue la référence d'axe ---
                .CodeAction = "FORCER_REF_LEVAGE"
                .LibelleAction = "Force la descente pour atteindre le capteur bas + référence codeur"
            
            Case DEBUT_TEMPO To FIN_TEMPO
                '--- temporisation ---
                .CodeAction = "TEMPO"
                .LibelleAction = "Temporisation de " & NumAction - DEBUT_TEMPO & " seconde(s)"
            
            Case TEMPO_EGOUTTAGE
                '--- temporisation d'égouttage ---
                .CodeAction = "TEMPO_EGOUTTAGE"
                .LibelleAction = "Temporisation d'égouttage"
            
            Case TEMPO_DEGRAISSAGE
                '--- temporisation de dégraissage ---
                .CodeAction = "TEMPO_DEGRAISSAGE"
                .LibelleAction = "Temporisation de dégraissage"
            
            Case TEMPO_RINCAGE
                '--- temporisation au rinçage ---
                .CodeAction = "TEMPO_RINCAGE"
                .LibelleAction = "Temporisation au rinçage"
            
            Case MONTEE_IMPULSION_CHARG
               '--- demande de MONTEE à un niveau conditionné par une impulsion pour le CHARGEMENT ---
                .CodeAction = "MONTEE_IMPULSION_CHARG"
                .LibelleAction = "Demande de MONTEE à un niveau conditionné par une impulsion pour le CHARGEMENT"
            
            Case DESCENTE_IMPULS_DECHARG
                '--- demande de DESCENTE à un niveau conditionné par une impulsion pour le DECHARGEMENT ---
                .CodeAction = "DESCENTE_IMPULS_DECHARG"
                .LibelleAction = "Demande de DESCENTE à un niveau conditionné par une impulsion pour le DECHARGEMENT"
            
            Case ATTENTE_AUTOR_DEPL_DECH    'code 520
                '--- attente de l'autorisation de DEPLACEMENT au POSTE de déchargement ---
                .CodeAction = "DESCENTE_IMPULS_DECHARG"
                .LibelleAction = "Attente de l'autorisation de DEPLACEMENT au POSTE de déchargement"
            
            Case SORTIE_BACS_ANTI_EGOUT
                '--- sortie des bacs anti-égouttures ---
                .CodeAction = "SORTIE_BACS_ANTI_EGOUT"
                .LibelleAction = "SORTIE des bacs anti-égouttures"
    
            Case GARAGE_BACS_ANTI_EGOUT
                '--- position garage des bacs anti-égouttures ---
                .CodeAction = "GARAGE_BACS_ANTI_EGOUT"
                .LibelleAction = "Position GARAGE des bacs anti-égouttures"
    
            Case CTRL_BACS_ANTI_EGOUT
                '--- contrôle de la position garage des bacs anti-égouttures ---
                .CodeAction = "CTRL_BACS_ANTI_EGOUT"
                .LibelleAction = "Contrôle de la position GARAGE des bacs anti-égouttures"
         
            Case SYNCHRO_CHARGEMENT_AUTO
                '--- SYNCHRO avec le CHARGEMENT AUTOMATIQUE ---
                .CodeAction = "SYNCHRO_CHARGEMENT_AUTO"
                .LibelleAction = "SYNCHRO avec le CHARGEMENT AUTOMATIQUE"
            
            Case SYNCHRO_CHARGEUR_AUTO
                '--- SYNCHRO avec le CHARGEUR AUTOMATIQUE MODE NORMAL, demande de mise en position du chargeur ---
                .CodeAction = "SYNCHRO_CHARGEUR_EN_POSITION"
                .LibelleAction = "SYNCHRO avec le CHARGEUR en AUTOMATIQUE, mise en position de Chargement"
             
            Case APPEL_NIVEAUX_CHARGEMENT
                '--- appel de la table des niveaux du POSTE de CHARGEMENT ---
                .CodeAction = "APPEL_NIVEAUX_CHARGEMENT"
                .LibelleAction = "Appel de la table des niveaux du POSTE de CHARGEMENT"
             
            Case APPEL_NIVEAUX_DECHARG
                '--- appel de la table des niveaux du POSTE de DECHARGEMENT ---
                .CodeAction = "APPEL_NIVEAUX_DECHARG"
                .LibelleAction = "Appel de la table des niveaux du POSTE de DECHARGEMENT"

            Case RAZ_CHARGE
                '--- RAZ de la CHARGE ---
                .CodeAction = "RAZ_CHARGE"
                .LibelleAction = "RAZ de la CHARGE"

            Case ACQUIS_OF
                '--- acquisition de l'ordre de fabrication ---
                .CodeAction = "ACQUIS_OF"
                .LibelleAction = "Acquisition de l'ordre de fabrication"

            Case AFFECT_CHARGE_PONT
                '--- affectation du numéro de charge sur le pont au chargement ---
                .CodeAction = "AFFECT_CHARGE_PONT"
                .LibelleAction = "Affectation du numéro de charge sur le pont au chargement"
            
            Case ATTENTE_VALID_APRES_CHAR
                '--- attente de la validation après le chargement ---
                .CodeAction = "ATTENTE_VALID_APRES_CHAR"
                .LibelleAction = "Attente de la validation après le chargement"
            
            Case ATTENTE_DESCENTE_DEGRAIS
                '--- attente pour la descente au poste de dégraissage ---
                .CodeAction = "ATTENTE_DESCENTE_DEGRAIS"
                .LibelleAction = "Attente pour la descente au poste de DEGRAISSAGE"
    
            Case ATTENTE_AU_DESSUS_CHROME
                '--- attente au dessus du chromage si le pont est arrivée trop tôt ---
                .CodeAction = "ATTENTE_AU_DESSUS_CHROME"
                .LibelleAction = "Attente au dessus du chromage si le pont est arrivé trop tôt"

            Case ATTENTE_DESCENTE_ATTAQUE
                '--- attente pour la descente au poste d'attaque ---
                .CodeAction = "ATTENTE_DESCENTE_ATTAQUE"
                .LibelleAction = "Attente pour la descente au poste d'ATTAQUE"
            
            Case CTRL_COMMUT_PASS_DEGRAIS
                '--- contrôle du commutateur de passage au DEGRAISSAGE (code 2005) ---
                .CodeAction = "CTRL_COMMUT_PASS_DEGRAIS"
                .LibelleAction = "Contrôle du commutateur de passage au DEGRAISSAGE"
    
            Case CTRL_COMMUT_PASS_ATTAQUE
                '--- contrôle du commutateur de passage à l'ATTAQUE (code 2006)---
                .CodeAction = "CTRL_COMMUT_PASS_ATTAQUE"
                .LibelleAction = "Contrôle du commutateur de passage à l'ATTAQUE"

            Case ATTENTE_FIN_CHROMAGE
                 '--- attente fin du cycle de chromage pour descente au poste ---
                .CodeAction = "ATTENTE_FIN_CHROMAGE"
                .LibelleAction = "Attente de la fin du cycle de chromage pour descendre au poste"
            
            Case ATTENTE_ARRET_REDRESSEUR
                '--- attente de l'arrêt d'un redresseur (contrôle sur l'intensité) ---
                .CodeAction = "ATTENTE_ARRET_REDRESSEUR"
                .LibelleAction = "Attente de l'arrêt d'un redresseur (contrôle sur l'intensité)"
            
            Case CRTL_SUIVI_AV_DESC
                '--- contrôle du suivi avant descente à un poste ---
                .CodeAction = "CRTL_SUIVI_AV_DESC"
                .LibelleAction = "CONTROLE du SUIVI AVANT DESCENTE  à un POSTE"
            
            Case CTRL_DEM_CHANGE_OUTIL
                '--- contrôle de la demande de changement outil ---
                .CodeAction = "CTRL_DEM_CHANGE_OUTIL"
                .LibelleAction = "Contrôle de la demande de changement outil"
    
            Case DEM_SORTIE_TRANSFERT
                '--- demande de SORTIE du chariot de transfert (EN LIGNE) ---
                .CodeAction = "DEM_SORTIE_TRANSFERT"
                .LibelleAction = "Demande de SORTIE du chariot de transfert (EN LIGNE)"
    
            Case DEM_RENTREE_TRANSFERT
                '--- demande de RENTREE du chariot de transfert (HORS LIGNE) ---
                .CodeAction = "DEM_RENTREE_TRANSFERT"
                .LibelleAction = "Demande de RENTREE du chariot de transfert (HORS LIGNE)"
    
            Case DEM_BON_NIVEAU_TRANFERT
                '--- demande de mise au bon niveau du chariot de transfert ---
                .CodeAction = "DEM_BON_NIVEAU_TRANFERT"
                .LibelleAction = "Demande de mise au bon niveau du chariot de transfert"
            
            Case CTRL_TRANSFERT_SORTIE
                '--- contrôle chariot de transfert SORTIE (EN LIGNE) ---
                .CodeAction = "CTRL_TRANSFERT_SORTIE"
                .LibelleAction = "Contrôle chariot de transfert SORTIE (EN LIGNE)"
    
            Case CTRL_TRANSFERT_RENTRE
                '--- contrôle chariot de transfert RENTRE (HORS LIGNE) ---
                .CodeAction = "CTRL_TRANSFERT_RENTRE"
                .LibelleAction = "Contrôle chariot de transfert RENTRE (HORS LIGNE)"
            
            Case CTRL_BON_NIV_TRANSFERT
                '--- contrôle du bon niveau du chariot de transfert ---
                .CodeAction = "CTRL_BON_NIV_TRANSFERT"
                .LibelleAction = "Contrôle du bon niveau du chariot de transfert"
            
            Case PRISE_DEPOSE_TRANSFERT
                '--- affectation des postes de dépose et de prise pour le transfert ---
                .CodeAction = "PRISE_DEPOSE_TRANSFERT"
                .LibelleAction = "Affectation des postes de dépose et de prise pour le TRANSFERT"
    
            Case TRL_POSTE_DEPOSE_TRANS
                '--- translation au poste de DEPOSE sur le chariot de transfert (emplacement vide) ---
                .CodeAction = "TRL_POSTE_DEPOSE_TRANS"
                .LibelleAction = "Translation au poste de DEPOSE sur le chariot de TRANSFERT (emplacement vide)"
    
            Case TRL_POSTE_PRISE_TRANS
                '--- translation au poste de PRISE sur le chariot de transfert ---
                .CodeAction = "TRL_POSTE_PRISE_TRANS"
                .LibelleAction = "Translation au poste de PRISE sur le chariot de TRANSFERT"

            Case ATTENTE_VALID_OUTIL
                '--- attente du code de validation du nouvel outil ---
                .CodeAction = "ATTENTE_VALID_OUTIL"
                .LibelleAction = "Attente du code de validation du nouvel outil"

            Case CTRL_FORCER_MANUEL
                '--- contrôle de forçage en manuel du pont ---
                .CodeAction = "CTRL_FORCER_MANUEL"
                .LibelleAction = "Contrôle de forçage en manuel du pont"

            Case CTRL_DESCENTE_DEGRAIS
                '--- contrôle si la descente est possible (poste 1) du dégraissage ---
                .CodeAction = "CTRL_DESCENTE_DEGRAIS"
                .LibelleAction = "Contrôle si la descente au premier poste du DEGRAISSAGE est possible"
            
            Case DEM_DEMI_POSTE_DEGRAIS
                '--- demande d'avance au demi poste du dégraissage ---
                .CodeAction = "DEM_DEMI_POSTE_DEGRAIS"
                .LibelleAction = "Demande d'avance au demi poste du dégraissage"

            Case LANCEMENT_TEMPO_DEGRAIS
                '--- lancement de la temporisation de dégraissage ---
                .CodeAction = "LANCEMENT_TEMPO_DEGRAIS"
                .LibelleAction = "Lancement de la temporisation de dégraissage"
            
            Case ARRET_TRL_DEGRAISSAGE
                '--- arrêt des 2 translateurs du dégraissage ---
                .CodeAction = "ARRET_TRL_DEGRAISSAGE"
                .LibelleAction = "Arrêt des 2 translateurs du dégraissage"
    
            Case REF_AXE_TRL_DEGRAISSAGE
                '--- référence d'axes des 2 translateurs du dégraissage ---
                .CodeAction = "REF_AXE_TRL_DEGRAISSAGE"
                .LibelleAction = "Référence d'axes des 2 translateurs du dégraissage"
    
            Case CTRL_DESCENTE_ATTAQ
                '--- contrôle si la descente est possible à l'attaque ---
                .CodeAction = " CTRL_DESCENTE_ATTAQ"
                .LibelleAction = "Contrôle si la descente est possible à l'ATTAQUE"

            Case ATTENTE_FIN_ATTAQUE
                '--- attente de la fin de l'attaque ---
                .CodeAction = "ATTENTE_FIN_ATTAQUE"
                .LibelleAction = "Attente de la fin de l'attaque"
    
            Case AUTOR_ATTAQUE_MODE_2P
                '--- autorisation de lancement de l'attaque en mode 2 ponts ---
                .CodeAction = "AUTOR_ATTAQUE_MODE_2P"
                .LibelleAction = "Autorisation de lancement de l'attaque en mode 2 ponts"
    
            Case CTRL_DESCENTE_REPRISE
                '--- contrôle si la descente est possible au poste de reprise ---
                .CodeAction = " CTRL_DESCENTE_REPRISE"
                .LibelleAction = "Contrôle si la descente est possible au POSTE de REPRISE"
    
            Case CHOIX_CYCLE
                '--- choix du cycle (chromage ou dégraissage) ---
                .CodeAction = "CHOIX_CYCLE"
                .LibelleAction = "Choix du cycle (chromage ou dégraissage)"
            
            Case CHOIX_MODE_CHARGEMENT
                '--- Choix du mode de chargement avec INIT (Polisseuse, Forcer Listo OU Chargeur), code ---
                '--- Aiguillage +1 Mode chargeur automatique, Aiguillage +2 Autres mode
                .CodeAction = "CHOIX_MODE_CHARGEMENT"
                .LibelleAction = "Choix du mode de Chargement (Chargeur AUTO=Pointeur+1, Polisseuse, Listo=Pointeur+2"
                
            Case LANCE_DECH_AUTO
                '--- lancement du déchargement automatique ---
                .CodeAction = "LANCE_DECH_AUTO"
                .LibelleAction = "Lancement du déchargement automatique"
    
            Case DEBUT_SYNCHRO_P1 To FIN_SYNCHRO_P1
                '--- synchronisation relatif au pont 1 (synchros de 1 à 10) ---
                .CodeAction = "SYNCHROS_P1"
                .LibelleAction = "Synchronisation " & NumAction - DEBUT_SYNCHRO_P1 + 1 & " du PONT 1"
            
            Case DEBUT_SYNCHRO_P2 To FIN_SYNCHRO_P2
                '--- synchronisation relatif au pont 2 (synchros de 1 à 10) ---
                .CodeAction = "SYNCHROS_P2"
                .LibelleAction = "Synchronisation " & NumAction - DEBUT_SYNCHRO_P2 + 1 & " du PONT 2"

            Case FCY
                '--- fin de cycle ---
               .CodeAction = "FCY"
               .LibelleAction = "Fin de cycle"
           
            Case DEBUT_SEQ_SAUT To FIN_SEQ_SAUT
                '--- séquence de saut (équivalent GOTO) ---
                .CodeAction = "SEQUENCE_SAUT"
                .LibelleAction = "Saut en " & NumAction - DEBUT_SEQ_SAUT
           
            Case Else
                '--- action inconnue ---
                .CodeAction = "INCONNUE"
                .LibelleAction = "INCONNUE"
        
        End Select
    End With

End Function


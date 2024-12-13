Attribute VB_Name = "MPConstantes"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES CONSTANTES PUBLIQUES
' Nom                    : MPConstantes.bas
' Date de création : 14/10/2010
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'*** ENUMERATIONS ***

'--- zones de déplacements des fenêtres ---
Public Enum ZONES_DEPLACEMENT_FENETRE
    Z_MERE = 0
    Z_FILLE = 1
End Enum

'--- repérage pour appels en divers endroits des fenêtres ---
Public Enum FENETRES
    
    F_SYNOPTIQUE = 2
    
    F_ORGANISATION_LIGNE = 100
    F_MOTEUR_INFERENCE = 101
    F_MODE_CYCLIQUE = 102
    
    F_GAMMES_ANODISATION = 200
    F_TRACABILITE_PRODUCTION = 201
    F_CHARGES_EN_LIGNE = 202

    F_CYCLES_PONTS = 204
    F_CHARGEMENT_PREVISIONNEL = 205
    F_GESTION_REDRESSEURS = 206
    F_GESTION_CUVES = 207
    F_GESTION_REGULATION = 208
    F_PROGRAMMATEUR_CYCLIQUE = 209
    F_ANNEXES = 210
    F_LISTE_DEFAUTS = 211

    F_PREMISSES = 400
    F_TEMPS_MOUVEMENTS = 401
    F_TRACABILITE_ALARMES = 402
    
    F_MAINTENANCE = 500
    F_INFORMATIONS_DEFAUTS_VARIATEURS = 501
    F_INFORMATIONS_DEFAUTS_COMMUNICATION_AUTOMATE = 502
    
    F_ESSAIS = 600

    F_APROPOS = 1300
    
    F_VISUALISATION_GRAPHES_PRODUCTION = 2010
    F_NETTOYAGE_GRAPHES_PRODUCTION = 2011
    
    F_MODIFICATION_OPTIONS_CHARGE = 3000
    
    F_MESSAGE = 3010
    
    F_MODIFICATIONS_AVANT_IMPRESSION = 4000
    F_CHOIX_IMPRESSION = 4010
    
    DR_GAMMES_ANODISATION = 10200
    DR_TRACABILITE = 10210
    DR_ALARMES_LIGNE = 10220
    
End Enum

'--- modes des outils du menu principal ---
Public Enum OUTILS_MENU_PRINCIPAL
    O_STANDARD = 0                                           'outils standard
    O_MODE_IA_CYCLIQUE = 1                           'outils pour la gestion du mode I.A. et du mode cyclique
    O_PRODUCTION = 2                                       'outils de production
End Enum

'--- type de boutons pour les barres d'outils ---
Public Enum TYPES_BOUTONS_OUTILS

    B_VIDE = 0                                                     'vide le bouton complétement et supprime le séparateur qui suit
    B_SEPARATEUR = 1                                      'installe un séparateur (barre de séparation)
    
    B_APERCU_AVANT_IMPRESSION = 11         'aperçu avant impression
    B_CALCULATRICE = 12                                  'calculatrice

    B_ORGANISATION_LIGNE = 21                      'organisation de la ligne
    B_MOTEUR_INFERENCE = 22                       'moteur d'inférence
    B_MODE_CYCLIQUE = 23                              'mode cyclique

    B_GAMMES_PRODUCTION = 31                   'gammes de production
    B_TRACABILITE_PRODUCTION = 32             'traçabilité de production
    B_CHARGES_EN_LIGNE = 33                        'charges en ligne
    B_CYCLES_PONTS = 34                                'cycles des ponts
    B_CHARGEMENT_PREVISIONNEL = 35        'chargement / prévisionnel
    B_REDRESSEURS = 36                                 'gestion des redresseurs
    B_CUVES = 37                                                'gestion des cuves
    B_REGULATION = 38                                     'gestion de la régulation
    B_PROGRAMMATEUR_CYCLIQUE = 39        'programmateur cyclique
    B_ANNEXES = 40                                           'annexes
    B_LISTE_DEFAUTS = 41                                'liste des défauts
    B_MAINTENANCE = 42                                   'maintenance
    B_FERMER_TOUT = 43                                  'fermeture de toutes les fenêtres

End Enum

'--- ponts sur la ligne ---
Public Enum PONTS
    P_1 = 1
    P_2 = 2
End Enum

'--- modes des ponts ---
Public Enum MODES_PONTS
    M_MAINTENANCE = 1
    M_MANUEL = 2
    M_SEMI_AUTOMATIQUE = 3
    M_AUTOMATIQUE = 4
End Enum

'--- types de séquences des ponts ---
Public Enum TYPES_SEQUENCES
    TS_INCONNU = 0
    TS_CYCLIQUE_PAR_IMPULSIONS = 1
    TS_CYCLIQUE_OPTIMISE = 2
    TS_ALEATOIRE = 3
End Enum

'--- types de cycles des ponts ---
Public Enum TYPES_CYCLES
    TC_INCONNU = 0                                      'type de cycle inconnu
    TC_DEPLACEMENT_PONT = 1                 'déplacement du pont (pour positionnement à un poste)
    TC_TRANSFERT_CHARGE = 2                 'transfert d'une charge d'un poste à un autre
End Enum

'--- niveaux sur les ponts ---
Public Enum NIVEAUX_PONTS
    N_BAS = 1
    N_INTERMEDIAIRE = 2
    N_HAUT = 15
End Enum

'--- sens X (translation) pour les ponts ---
Public Enum SENS_X
    S_AVANT = 1
    S_ARRIERE = -1
    S_AU_POSTE = 0
End Enum
    
'--- sens Y (montée/descente) pour les ponts ---
Public Enum SENS_Y
    S_MONTEE = 1
    S_DESCENTE = -1
    S_AU_NIVEAU = 0
End Enum

'--- 0=accroches de la charge en haut, 1=accroches de la charge en bas ---
Public Enum ETATS_ACCROCHES
    E_ACCROCHES_EN_HAUT = 0
    E_ACCROCHES_EN_BAS = 1
    E_ACCROCHES_VERS_HAUT = 2
    E_ACCROCHES_VERS_BAS = 3
End Enum

'--- postes de la ligne d'anodisation ---
Public Enum POSTES
    
    P_CHGT_1 = 1          'chargement 1
    P_CHGT_2 = 2          'chargement 2
    'P_CHGT_3 = 3          'chargement 3
    P_C02 = 4                 'réserve
    P_C00 = 5             'chargement 4

    P_DEC = 6                 'décapage
    P_SAT = 7                 'satinage S201
    
    P_C03 = 8                 'futur décapage
    P_C04 = 9                 'rinçage dégraissage
    
    P_C05 = 10                'brillantage n°1
    P_C06 = 11                'rinçage Mt brillantage
    P_C07 = 12                'brillantage n°2
    P_C08 = 13                'rinçage brillantage
    P_C09 = 14                'rinçage brillantage
    
    P_C10 = 15                'neutralisation
    P_C11 = 16                'rinçage blanchiment
    P_C12 = 17                'blanchiment
    
    P_C13 = 18                'anodisation
    P_C14 = 19                'anodisation
    P_C15 = 20                'anodisation
    P_C16 = 21                'anodisation
    P_C17 = 22                'rinçage anodisation
    P_C18 = 23                'rinçage anodisation

    P_C19 = 24                'spectrocoloration
    P_C20 = 25                'rinçage
    P_C21 = 26                'rinçage
    P_C22 = 27                'coloration or
    
    P_C23 = 28                'RESERVE 1
    P_C24 = 29                'RESERVE 2
    P_C25 = 30                'RESERVE 3
    P_C26 = 31                'RESERVE 4

    P_C27 = 32                'imprégnation à froid
    P_C28 = 33                'coloration noire
    P_C29 = 34                'rinçage noir
    P_C30 = 35                'rinçage eau dure/imprégnation

    P_C31 = 36                'Imprégnation à froid
    
    P_C32 = 37                'colmatage chaud
    P_C33 = 38                'colmatage chaud
    
    P_C34 = 39                'séchoir poste
    
    P_C35 = 40                'RESERVE 5
    
    P_D1 = 41                  'déchargement 1
    P_D2 = 42                  'déchargement 2
    
   ' P_C36 = 433                '
    P_C37 = 43               '
    P_C38 = 44                '
End Enum

Public Const ZONE_ETUVE  As Integer = 37

Public Const PREMIER_POSTE  As Integer = 1
Public Const PREMIER_BAIN  As Integer = 4
Public Const DERNIER_POSTE  As Integer = 44

Public Const PREMIERE_CUVE  As Integer = 1
Public Const DERNIERE_CUVE  As Integer = 40




'--- cuves de la ligne d'anodisation ---
Public Enum CUVES
    
    C_C00 = 1                  'dégraissage
    C_DEC = 2                  'décapage
    C_SAT = 3                  'satinage
    
    
    C_C02 = 40                  'Reserve
    C_C03 = 4                  'rinçage soude
    C_C04 = 5                  'rinçage dégraissage
    
    C_C05 = 6                  'brillantage n°1
    C_C06 = 7                  'rinçage Mt brillantage
    C_C07 = 8                  'brillantage n°2
    C_C08 = 9                  'rinçage brillantage
    C_C09 = 10                'rinçage brillantage
    
    C_C10 = 11                'neutralisation
    C_C11 = 12                'rinçage blanchiment
    C_C12 = 13                'blanchiment
    
    C_C13 = 14                'anodisation
    C_C14 = 15                'anodisation
    C_C15 = 16                'anodisation
    C_C16 = 17                'anodisation
    C_C17 = 18                'rinçage anodisation
    C_C18 = 19                'rinçage anodisation

    C_C19 = 20                'spectrocoloration
    C_C20 = 21                'rinçage
    C_C21 = 22                'rinçage
    C_C22 = 23                'coloration or
    
    C_C23 = 24                'RESERVE 1
    C_C24 = 25                'RESERVE 2
    C_C25 = 26                'RESERVE 3
    C_C26 = 27                'RESERVE 4

    C_C27 = 28                'imprégnation à froid
    C_C28 = 29                'coloration noire
    C_C29 = 30                'rinçage noir
    C_C30 = 31                'rinçage eau dure/imprégnation

    C_C31 = 32                'colmatage chaud
    C_C32 = 33                'colmatage chaud
    
    C_C33 = 34                'Colmatage froid
    C_C34 = 35                'réserve
    C_C35 = 36                'réserve
    C_C36 = 37                'réserve
    C_C37 = 38                'étuve
    C_C38 = 39                'basculeur
    
End Enum
'--- cuves gérées par l'automate ---
Public Const PREMIERE_CUVE_API_OLD   As Integer = 1
Public Const DERNIERE_CUV_API_OLD  As Integer = 24
Public Enum CUVES_API_OLD
    
    C_C00 = 1                  'dégraissage
    C_DEC = 2                  'décapage
    C_SAT = 3                  'satinage S201
    C_C03 = 4                  'rinçage soude
    C_C05 = 5                  'brillantage n°1
    C_C06 = 6                  'rinçage Mt brillantage
    C_C07 = 7                  'brillantage n°2
    C_C13 = 8                  'anodisation
    C_C14 = 9                  'anodisation
    C_C15 = 10                'anodisation
    C_C16 = 11                'anodisation
    C_C19 = 12                'spectrocoloration
    C_C22 = 13                'coloration or
    C_C27 = 14                'imprégnation à froid
    C_C28 = 15                'coloration noire
    C_C31 = 16                'colmatage chaud
    C_C32 = 17                'colmatage chaud
    C_C33 = 18                'colmatage froid
    C_C34 = 19                'réserve
    C_C35 = 20                'réserve
    C_C36 = 21                'réserve
    C_C37 = 22                'étuve
    C_C38 = 23                'basculeur
    C_C02 = 24               'Réserve
   
    
End Enum


  

Public Const DERNIERE_CUV_REGULATION = 11
'--- cuves qui servent au formualire et au bon affichage ---
Public Enum CUVES_REGULATION
    C_C00 = 1                  'dégraissage
    C_DEC = 2                  'satinage S201
    C_C07 = 3                  'brillantage n°2 --> 7
    
    C_C13 = 4                  'anodisation   --> 8
    C_C14 = 5                  'anodisation   --> 9
    C_C15 = 6                  'anodisation   --> 10
   
  
    C_C22 = 7                'coloration or  -->  13
    C_C27 = 8                'imprégnation à froid --> 14
    C_C28 = 9                'coloration noire  -->  15
    C_C31 = 10                'colmatage chaud  --> 16
    C_C32 = 11                'colmatage chaud  --> 17
   

    
End Enum

'--- valeur codeur du niveau haut des ponts ---
Public Const VALEUR_CODEUR_NIVEAU_HAUT_PONTS As Integer = 2140

'--- cuves pour le déclenchement par les températures de la ventilation en mode automatique ---
Public Enum CUVES_TEMP_VENTILATION

    C_A1 = 1                    'pré-dégraissage 90°C
    C_A2 = 2                    'dégraissage 70°C
    C_A3 = 3                    'dégraissage électro.
    C_A7 = 4                    'décapage HCl 50%
    C_A8 = 5                    'décapage H2SO4 15%
    C_C13 = 6                  'première cuve d'anodisation chimique 93°C maxi.
    C_C14 = 7                  'deuxième cuve d'anodisation chimique 93°C maxi.
    C_A17 = 8                  'rinçage chaud déminé 75°C
    
    C_B1 = 9                    'dégraissage 70°C
    C_B4 = 10                  'activation alu. 20°C
    C_B5 = 11                  'activation HNO3 70%
    C_B8 = 12                  'zincate 25°C
    C_C15 = 13                'troisième cuve d'anodisation chimique 93°C maxi.

End Enum

'--- numéros de charges avec mini et maxi ---
Public Enum CHARGES
    PAS_DE_CHARGE = 0             'pas de charge (valeur zéro dans ce cas)
    C_NUM_MINI = 1                     'numéro de charge mini
    C_NUM_MAXI = 15                  'numéro de charge maxi
End Enum

'--- numéros de barres avec mini et maxi ---
Public Enum BARRES
    B_NUM_MINI = 1                     'numéro de barre mini
    B_NUM_MAXI = 57                  'numéro de barre maxi
End Enum

'--- choix du poste d'anodisation ---
Public Enum CHOIX_POSTE_ANODISATION
    C_AUTOMATIQUE = 0
    C_C13_IMPOSE = 1
    C_C14_IMPOSE = 2
    C_C15_IMPOSE = 3
    C_C16_IMPOSE = 4
End Enum

'--- cycles des ponts ---
Public Enum CYCLES
    C_ACTUEL = 0
    C_PROCHAIN = 1
End Enum

'--- types de PC ---
Public Enum TYPES_PC
    PC_SUR_LIGNE = 1                                    'ordinateur sur la ligne d'anodisation
    PC_ENTREPRISE = 2                                  'ordinateur quelconque de l'entreprise (connecté au réseau interne)
    PC_DISTANT = 3                                          'ordinateur distant (liaison par modem)
End Enum
'--- types ENV BDD ANO---
Public Enum TYPE_BDD_ANO
    PROD = 1
    TEST = 2
End Enum
Public Enum TYPE_BDD_CLIPPER
    HF_PROD = 1
    HF_TEST = 2
    ACCESS_TEST = 3
End Enum

'--- les défauts ---
Public Enum DEFAUTS
    NUM_MINI = 1
    NUM_MAXI = 1000
End Enum

'--- pour la gestion des grilles de données ---
Public Enum GESTION_GRILLES
    GG_INITIALISATION = 1
    GG_VIDAGE = 2
    GG_TRANSFERT_DONNEES = 3
    GG_COMPRESSION = 4
    GG_AFFICHAGE = 5
    GG_MEMORISATION = 6
End Enum

'--- types d'impressions ---
Public Enum TYPES_IMPRESSIONS
    TI_APERCU_AVANT_IMPRESSION = 0
    TI_IMPRIMER = 1
    TI_IMPRIMER_FENETRE_ACTIVE = 2
End Enum

'--- échelles 24 heures ---
Public Enum ECHELLES_24H
    E_CHAUFFAGE = 0
    E_POMPE_CHAUFFAGE = 1
    E_VENTILATION_CHAUFFAGE = 2
End Enum

'--- position des indicateurs ---
Public Enum INDICATEURS
    I_VERT = 0
    I_ORANGE = 1
    I_ROUGE = 2
    I_PETIT_VERT = 3
    I_PETIT_ROUGE = 4
End Enum

'--- position des flèches ---
Public Enum FLECHES
    F_VERTE = 0
    F_ORANGE = 1
    F_ROUGE = 2
End Enum

'--- bases de données ---
Public Enum BASES_DONNEES
    BD_ANODISATION_SQL = 0
    BD_SAGE_SQL = 1
End Enum

'--- mots de passe ---
Public Enum MOTS_DE_PASSE
    MDP_DIRECTION = 0
    MDP_PERSONNEL = 1
End Enum

'--- marges pour l'affichage des boutons, etc... ---
Public Enum MARGES
    M_BORD_GAUCHE = 120
    M_BORD_DROIT = 120
    M_BORD_HAUT = 120
    M_BORD_BAS = 140
    M_ENTRE_BOUTONS = 140
    M_BORS_BAS_GRILLE = 160
End Enum

'--- types de boutons pour les enregistrements ---
Public Enum BOUTONS_ENREGISTREMENTS
    B_PREMIER = 0
    B_PRECEDENT = 1
    B_SUIVANT = 2
    B_DERNIER = 3
End Enum

'--- états des boutons des enregistrements ---
Public Enum ETATS_BOUTONS_ENREGISTREMENTS
    E_TOUT_INVISIBLE = 0
    E_TOUT_VISIBLE = 1
    E_PRECEDENT_SUIVANT = 2
    E_PREMIER_DERNIER = 3
End Enum

'--- états des boutons liés aux fiches ---
Public Enum ETATS_BOUTONS
    
    E_CHARGEMENT_FENETRE = 0
    E_DECHARGEMENT_FENETRE = 1
    
    E_MODIFICATION_EN_COURS = 2
    
    E_AVANT_DEPLACEMENT = 3
    E_AVANT_QUITTER = 4
    E_AVANT_VALIDER = 5
    E_AVANT_ANNULER = 6
    E_AVANT_RETABLIR = 7
    E_AVANT_ACTUALISER = 8
    E_AVANT_NOUVEAU = 9
    E_AVANT_SUPPRIMER = 10
    
    E_APRES_DEPLACEMENT = 11
    E_APRES_QUITTER = 12
    E_APRES_VALIDER = 13
    E_APRES_ANNULER = 14
    E_APRES_RETABLIR = 15
    E_APRES_ACTUALISER = 16
    E_APRES_NOUVEAU = 17
    E_APRES_SUPPRIMER = 18

End Enum

'--- formats des données ---
Public Enum DONNEES
    
    D_GENERALE = 1                                              'tous les caractères sans modification
    D_GENERALE_MINUSCULES = 2                      'tous les caractères en minuscules
    D_GENERALE_MAJUSCULES = 3                      'tous les caractères en majuscules
    
    D_TEXTE = 10                                                    'lettres de a à z en minuscules et ou majuscules
    D_TEXTE_MINUSCULES = 11                            'lettres de a à z en minuscules
    D_TEXTE_MINUSCULES_NUMERIQUES = 12   'lettres de a à z en minuscules et ou touches numériques
    D_TEXTE_MAJUSCULES = 13                            'lettres de a à z en majuscules
    D_TEXTE_MAJUSCULES_NUMERIQUES = 14   'lettres de a à z en majuscules et ou touches numériques
    
    D_NBR_NATURELS = 20                                    'touches numériques sans décimale positif (de 0 à x)
    D_NBR_RELATIFS = 21                                      'touches numériques sans décimale positif (de -x à +x)
    D_NBR_REELS = 22                                           'touches numériques avec décimale (de -x,x... à + x,xx...)
    D_NBR_REELS_POSITIFS = 23                          'touches numériques avec décimale (de 0 à + x,xx...)
    
    D_HEURE_HHMM = 30                                      'format heure HH:MM
    D_HEURE_HHMMSS = 31                                  'format heure HH:MM:SS
    
    D_DATE_JJMMAAAA = 40                                   'format date JJ/MM/AAAA
    
    D_TELEPHONE = 100                                         'format téléphone (03-10-20-24-26)
    D_FAX = 101                                                       'format fax (03-10-20-24-26)
    
    D_CODE_CLIENT = 199                                      'format code client
    D_CODE_FOURNISSEUR = 200                         'format code fournisseur
    D_TYPE_DE_PRIX = 201                                    'format type de prix (U ou E en majuscules)
    D_AVEC_JUMELAGE = 202                                 'format avec jumelage (Espace ou D (double))
    D_JOUR_OU_NUIT = 203                                   'format nuit ou jour (J ou N en majuscules)
    D_MANU_AUTO = 204                                         'format manu auto (A ou M en majuscules)
    
    D_CODE_POSTAL = 300                                     'format code postal

    D_SECURITE_SOCIALE = 400                            'sécurité sociale

End Enum

'--- couleurs pour la construction des échelles graphiques ---
Public Enum COULEURS_ECHELLES_GRAPHIQUES
    C_ARRET_POMPE = COULEURS.BLANC
    C_MARCHE_POMPE = COULEURS.ORANGE_3
    C_MODE_ARRET = COULEURS.BLANC
    C_MODE_VEILLE = COULEURS.CYAN_3
    C_MODE_PRODUCTION = COULEURS.ORANGE_3
End Enum

'--- journées types ---
Public Enum JOURNEES_TYPES
    J_ARRET = 0
    J_TRAVAIL = 1
    J_VEILLE = 2
    J_REPRISE = 3
End Enum

'--- modes de production (couleurs sur l'échelle graphique) ---
Public Enum MODES_PRODUCTION
    M_ARRET = 0
    M_VEILLE = 1
    M_PRODUCTION = 2
End Enum

'--- modes affichages du synoptique ---
Public Enum MODES_AFFICHAGES_SYNOPTIQUE
    MA_NUM_CHARGES = 0                                   'affichage avec les numéros des charges
    MA_NUM_BARRES = 1                                     'affichage avec les numéros des barres
    MA_COLORATIONS = 2                                     'affichage des charges passants en coloration
End Enum
    
'***************************************************************************************************************************
'                                                        ELEMENTS PHYSIQUES DE LA LIGNE
'***************************************************************************************************************************

'--- états des mouvements ---
Public Enum ETATS_MOUVEMENTS
    E_PAS_DE_MOUVEMENT = 0
    E_MOUVEMENT_EN_COURS = 1
    E_FIN_DU_MOUVEMENT = 2
End Enum

'--- états d'un chauffage ---
Public Enum ETATS_CHAUFFAGES
    M_ARRET = 0                           'chauffage à l'arrêt
    M_MARCHE = 1                        'chauffage en marche
    M_DEFAUT = 2                         'chauffage est en défaut
End Enum

'--- états du refroidissement d'un bain ---
Public Enum ETATS_REFROIDISSEMENT_BAIN
    M_ARRET = 0                           'chauffage à l'arrêt
    M_MARCHE = 1                        'chauffage en marche
    M_DEFAUT = 2                         'chauffage est en défaut
End Enum

'--- cas général pour la pompe ---
Public Enum MODES_POMPES
    M_FORCER_ARRET = 0
    M_FORCER_MARCHE = 1
    M_AUTO = 2
End Enum
Public Enum CYCLES_POMPES  'pour le mode automatique (MODES_POMPE=2)
    CP_ARRET = 0
    CP_MARCHE = 1
End Enum
Public Enum ETATS_POMPES     'états logique d'une pompe
    E_ARRET = 0
    E_MARCHE = 1
    E_DEFAUT = 2
End Enum

'--- cas général des couvercles ---
Public Enum MODES_COUVERCLES
    M_AUTO = 1
    M_FORCER_FERMETURE = 2
    M_FORCER_OUVERTURE = 3
End Enum
Public Enum CYCLES_COUVERCLES
    C_DEMANDE_FERMETURE = 0
    C_DEMANDE_OUVERTURE = 1
End Enum
Public Enum ETATS_COUVERCLES
    E_COUVERCLES_INDETERMINES = 0
    E_COUVERCLES_OUVERTS = 1
    E_COUVERCLES_FERMES = 2
    E_COUVERCLES_EN_OUVERTURE = 3
    E_COUVERCLES_EN_FERMETURE = 4
    E_DEFAUT_COUVERCLES = 5
End Enum

'--- cas général des niveaux des cuves ---
Public Enum ETATS_NIVEAUX
    E_BAS_X_MINUTES = 0
    E_TRES_BAS = 1
    E_INTERMEDIAIRE_BAS = 2
    E_NORMAL = 3
    E_INTERMEDIAIRE_HAUT = 4
    E_TRES_HAUT = 5
End Enum

'--- cas général de l'électro-vanne d'arrivée d'eau ---
Public Enum ETATS_EV_EAU
    E_FERMEE = 0
    E_OUVERTE = 1
    E_DEFAUT = 2
    E_DELAI_LONG = 3
End Enum

'--- états de la pompe de relevage de l'anodisation ---
Public Enum ETATS_POMPE_ANODISATION
    E_PAS_DE_DEFAUT = 0
    E_DEFAUT = 1
End Enum

'--- états du surpresseur d'eau ---
Public Enum ETATS_SURPRESSEUR_EAU
    E_PAS_DE_DEFAUT = 0
    E_DEFAUT = 1
End Enum

'--- modes de l'électro-vanne d'arrivée d'eau de la ligne ---
Public Enum MODES_EV_EAU_LIGNE
    M_ARRET = 0
    M_MARCHE = 1
End Enum
Public Enum ETATS_EV_EAU_LIGNE
    E_ARRET = 0
    E_MARCHE = 1
End Enum

'--- modes des compresseurs des ponts ---
Public Enum MODES_COMPRESSEURS_PONTS
    M_ARRET = 0
    M_MARCHE = 1
End Enum
Public Enum ETATS_COMPRESSEURS_PONTS
    E_ARRET = 0
    E_MARCHE = 1
End Enum

'--- modes de l'éclairage des ponts ---
Public Enum MODES_ECLAIRAGE_PONTS
    M_ARRET = 0
    M_MARCHE = 1
End Enum
Public Enum ETATS_ECLAIRAGE_PONTS
    E_ARRET = 0
    E_MARCHE = 1
End Enum

'--- modes de l'électro-vanne d'air dans le bain de brillantage ---
Public Enum MODES_EV_BRILLANTAGE
    M_ARRET = 0
    M_MARCHE_EN_AUTOMATIQUE = 1
    M_MARCHE_FORCEE = 2
End Enum
Public Enum ETATS_EV_BRILLANTAGE
    E_ARRET = 0
    E_MARCHE = 1
    E_DEFAUT = 2
End Enum

'--- états des chariots ---
Public Enum ETATS_CHARIOTS
    E_ABSENT = 0
    E_PRESENT = 1
    E_PRESENT_VERROUILLE = 2
End Enum

'--- ensembles des niveaux des rétentions ---
Public Enum NIVEAUX_RETENTIONS
    NR_STOCKAGE_STATION = 0
    NR_LIGNE_ANODISATION = 1
    NR_TRAITEMENT_EAUX = 2
    NR_LAVEUR = 3
End Enum
Public Enum ETATS_NIVEAUX_RETENTIONS
    ENR_INCONNU = 0
    ENR_BON = 1
    ENR_EN_DETECTION = 2
    ENR_HAUT = 3
End Enum

'--- numéros des redresseurs / nombre mini/maxi ---
Public Enum REDRESSEURS
    
    R_C13 = 1
    R_C14 = 2
    R_C15 = 3
    R_C16 = 4
    R_C19 = 5

    R_NUM_MINI = 1
    R_NUM_MAXI = 5

End Enum

'--- limites des redresseurs ---
Public Enum LIMITES_REDRESSEURS
    LM_TENSION = 20                                     'tension maximale de 20V
    LM_INTENSITE = 3000                              'intensité maximale de 3000A
End Enum

'--- images des redresseurs ---
Public Enum IMAGES_REDRESSEURS
    I_BAS_REDRESSEUR_VERT = 0
    I_BAS_REDRESSEUR_ORANGE = 1
    I_BAS_REDRESSEUR_BLANC = 2
    I_BAS_REDRESSEUR_ROUGE = 3
    I_BAS_REDRESSEUR_EXCLUS = 4
End Enum

'--- cas général des redresseurs ---
Public Enum MODES_REDRESSEUR
    MR_NON_DEFINI = 0
    MR_MANUEL = 1
    MR_AUTOMATIQUE = 2
End Enum

'--- états d'un redressseur ---
Public Enum ETATS_REDRESSEUR
    ER_ARRET = 0
    ER_PARTIEL = 1               'marche partiel lorsque un des redresseurs est à l'arrêt
    ER_MARCHE = 2
    ER_DEFAUT = 3
    ER_EXCLUSION = 4
End Enum

'--- modes U ou I des redresseurs ---
Public Enum MODES_U_OU_I
    M_TENSION = 0                               'mode pour les gammes en tension
    M_INTENSITE = 1                            'mode pour les gammes en intensité
End Enum

Public Enum SENS_REDRESSEUR
    SR_NON_DEFINI = 0
    SR_ANODIQUE = 1
    SR_CATHODIQUE = 2
    SR_SPECTRO = 3
End Enum

'--- les différentes phases d'une gamme redresseur ---
Public Enum PHASES_GAMMES_REDRESSEURS
    PH_T1 = 1
    PH_T2 = 2
    PH_T3 = 3
    PH_T4 = 4
End Enum

'--- mode de la régulation ---
Public Enum MODES_REGULATION
    MR_MANUEL = 0
    MR_AUTOMATIQUE = 1
End Enum

'--- types de prémisses ---
Public Enum TYPES_PREMISSES
    TP_DECODEES = 0                         'types de prémisses décodées (ex : 101-141-104)
    TP_CODEES = 1                              'types de prémisses codées (ex : NB-FCR-NH)
End Enum

'--- options d'une gamme au lancement de la charge ---
Public Enum OPTIONS_GAMME
    
    O_FORCER_MONTEE_EN_TPV = 0                                     'forcer la montée d'une charge en très petite vitesse
    O_FORCER_MONTEE_EN_PV = 1                                       'forcer la montée d'une charge en petite vitesse
    O_FORCER_DESCENTE_EN_TPV = 2                                 'forcer la descente d'une charge en très petite vitesse
    O_FORCER_DESCENTE_EN_PV = 3                                   'forcer la descente d'une charge en petite vitesse

    O_ACTIVER_AIR_BRILLANTAGE = 0                                   'activer l'air dans le brillantage

End Enum

'--- tous les cas de types de collision ---
Public Enum TYPES_COLLISION
    
    AUCUN_RISQUE = 0    'pas de risque de collision
    
    RISQUE_DEM_P1_AR_P2_AV = 1 '        PONT 2             PONT 1        OU   PONT 2   A <------------- D
    RISQUE_DEM_P2_AV_P1_AR = 2 '  A <------------- D     D -------------> A                                     D -------------> A   PONT 1
    
    RISQUE_DEM_P1_AV_P2_AR = 3 '       PONT 2               PONT 1       OU   PONT 2   D -------------> A
    RISQUE_DEM_P2_AR_P1_AV = 4 '  D -------------> A     A <------------- D                                     A <------------- D   PONT 1
     
    RISQUE_DEM_P1_AV_P2_AV = 5 '       PONT 2              PONT 1        OU  PONT 2   A <------------- D
    RISQUE_DEM_P2_AV_P1_AV = 6 '  A <------------- D     A <------------- D                                     A <------------- D   PONT 1

    RISQUE_DEM_P1_AR_P2_AR = 7 '      PONT 2             PONT 1         OU  PONT 2   D -------------> A
    RISQUE_DEM_P2_AR_P1_AR = 8 ' D -------------> A     D -------------> A                                     D -------------> A   PONT 1

End Enum

'--- types de base de données pour extraire les n° de fiches d'atelier ou d'affaires ---
Public Enum TYPES_BD
    BD_CLIPPER = 0                     'base de données CLIPPER
    BD_SAGE = 1                          'base de données SAGE
End Enum

'*** CONSTANTES NUMERIQUES ***

'--- programme ---
Public Const PROGRAMME_AVEC_AUTOMATE As Boolean = True         'pour simplifié le développement
Public Const PROGRAMME_TERMINE As Boolean = True                       'pour simplifié le développement

'--- temps minimum de stabilisation a vide (sans charge, temps en secondes) ---
Public Const TEMPS_MINI_STABILISATION_A_VIDE As Integer = 1

'--- temps minimum de stabilisation avec une charge (temps en secondes) ---
Public Const TEMPS_MINI_STABILISATION_AVEC_CHARGE As Integer = 1

'--- nombre de matières maximales par gamme ---
Public Const NBR_MATIERES_MAXI_PAR_GAMME As Integer = 10

'--- nombres de lignes pour les grilles de données et le nombre d'enregistrements extraits des requêtes ---
Public Const NBR_LIGNES_DETAILS_GAMMES_PRODUCTION  As Integer = 50
Public Const NBR_LIGNES_DETAILS_PREMISSES  As Integer = 100
Public Const NBR_LIGNES_CYCLES_PONTS As Integer = 50
Public Const NBR_LIGNES_PREVISIONNEL As Integer = 30
Public Const NBR_LIGNES_DETAILS_CHARGES As Integer = 50            'doit correspondre avec le chargement
Public Const NBR_LIGNES_DETAILS_REFERENCES_CLIENT  As Integer = 18
Public Const NBR_LIGNES_DETAILS_FICHES_PRODUCTION  As Integer = 100
Public Const NBR_LIGNES_TRAVAUX  As Integer = 18

'--- bits pour le traitement automate ---
Public Const BIT_0 As Integer = 0
Public Const BIT_1 As Integer = 1
Public Const BIT_2 As Integer = 2
Public Const BIT_3 As Integer = 3
Public Const BIT_4 As Integer = 4
Public Const BIT_5 As Integer = 5
Public Const BIT_6 As Integer = 6
Public Const BIT_7 As Integer = 7
Public Const BIT_8 As Integer = 8
Public Const BIT_9 As Integer = 9
Public Const BIT_10 As Integer = 10
Public Const BIT_11 As Integer = 11
Public Const BIT_12 As Integer = 12
Public Const BIT_13 As Integer = 13
Public Const BIT_14 As Integer = 14
Public Const BIT_15 As Integer = 15

'--- relatif aux fichiers ---
Public Const TEMPS_VALIDITE_FICHIER As Integer = 1           'temps de validité d'un fichier aprés le lancement de son écriture

'--- couleurs spéciales ---
Public Const ROUGE_DEFAUT As Long = &HF0&
Public Const ORANGE_CUVE As Long = &H6FB7FF

'--- nombre de mots d'une fiche de suivi dans l'API ---
Public Const NBR_MOTS_FICHE_SUIVI_API As Integer = 2

'--- nombre de mots d'une fiche cuve dans l'API ---
Public Const NBR_MOTS_FICHE_CUVE_API As Integer = 15

'--- positions des bits utilisés pour les états des commutations ---
Public Const POS_BIT_MANUEL_P2 As Integer = 0
Public Const POS_BIT_SEMI_AUTOMATIQUE_P2 As Integer = 1
Public Const POS_BIT_AUTOMATIQUE_P2 As Integer = 2
Public Const POS_BIT_MAINTENANCE_P2 As Integer = 3
Public Const POS_BIT_MANUEL_P1 As Integer = 8
Public Const POS_BIT_SEMI_AUTOMATIQUE_P1 As Integer = 9
Public Const POS_BIT_AUTOMATIQUE_P1 As Integer = 10
Public Const POS_BIT_MAINTENANCE_P1 As Integer = 11

'--- positions des bits utilisés pour les états des sécurités ---
Public Const POS_BIT_FRONT_MONTANT_DEFAUTS As Integer = 0
Public Const POS_BIT_STOP_LIGNE As Integer = 4
Public Const POS_BIT_ARRET_URGENCE_P1 As Integer = 5
Public Const POS_BIT_ARRET_URGENCE_P2 As Integer = 6

Public Const POS_BIT_MARCHE_GENERALE As Integer = 8
Public Const POS_BIT_ARRET_URGENCE As Integer = 9
Public Const POS_BIT_PORTILLONS_LIGNE_VIE As Integer = 10
Public Const POS_BIT_SECURITE_P1 As Integer = 11
Public Const POS_BIT_SECURITE_P2 As Integer = 12
Public Const POS_BIT_MANQUE_TENSION As Integer = 13
Public Const POS_BIT_MANQUE_AIR As Integer = 14
Public Const POS_BIT_ACQUITTEMENT_DEFAUTS As Integer = 15

'--- positions des bits utilisés pour les états 1 des cuves ---
Public Const POS_BIT_OUVERTURE_COUVERCLES As Integer = 0
Public Const POS_BIT_DEM_OUVERTURE_COUVERCLES As Integer = 1
Public Const POS_BIT_FERMETURE_COUVERCLES As Integer = 2
Public Const POS_BIT_DEM_FERMETURE_COUVERCLES As Integer = 3
Public Const POS_BIT_AGITATION_BAIN As Integer = 4
Public Const POS_BIT_DEM_AGITATION_BAIN As Integer = 5
Public Const POS_BIT_COUVERCLES_OUVERTS As Integer = 6
Public Const POS_BIT_COUVERCLES_FERMES As Integer = 7

Public Const POS_BIT_CHAUFFAGE As Integer = 8
Public Const POS_BIT_DEM_CHAUFFAGE As Integer = 9
Public Const POS_BIT_REFROIDISSEMENT As Integer = 10
Public Const POS_BIT_DEM_REFROIDISSEMENT As Integer = 11
Public Const POS_BIT_POMPE As Integer = 12
Public Const POS_BIT_DEM_POMPE As Integer = 13
Public Const POS_BIT_EV_EAU As Integer = 14
Public Const POS_BIT_DEM_EV_EAU As Integer = 15

'--- positions des bits utilisés pour les états 2 des cuves ---
Public Const POS_BIT_DEFAUT_CHAUFFAGE As Integer = 0
Public Const POS_BIT_DEFAUT_REFROIDISSEMENT As Integer = 1
Public Const POS_BIT_DEFAUT_POMPE As Integer = 2
Public Const POS_BIT_DEFAUT_EV_EAU As Integer = 3
Public Const POS_BIT_DEFAUT_COUVERCLES As Integer = 4
Public Const POS_BIT_DEFAUT_AGITATION_BAIN As Integer = 5

Public Const POS_BIT_NIVEAU_TRES_BAS As Integer = 8
Public Const POS_BIT_NIVEAU_INTERMEDIAIRE_BAS As Integer = 9
Public Const POS_BIT_NIVEAU_INTERMEDIAIRE_HAUT As Integer = 10
Public Const POS_BIT_NIVEAU_TRES_HAUT As Integer = 11
Public Const POS_BIT_MANU_AUTO_REGULATION As Integer = 12

'--- positions des bits utilisés pour les délais trop long des électro-vannes d'arrivée d'eau ---
'Public Const POS_BIT_DELAI_LONG_EV_A1 As Integer = 8
'Public Const POS_BIT_DELAI_LONG_EV_A2 As Integer = 9
'Public Const POS_BIT_DELAI_LONG_EV_A3 As Integer = 10
'Public Const POS_BIT_DELAI_LONG_EV_C13 As Integer = 11
'Public Const POS_BIT_DELAI_LONG_EV_C14 As Integer = 12
'Public Const POS_BIT_DELAI_LONG_EV_A17 As Integer = 13
'Public Const POS_BIT_DELAI_LONG_EV_B1 As Integer = 14
'Public Const POS_BIT_DELAI_LONG_EV_B8 As Integer = 15
'Public Const POS_BIT_DELAI_LONG_EV_C15 As Integer = 1

'--- positions des bits utilisés pour le chargement/déchargement ---
Public Const POS_BIT_D1 As Integer = 0
Public Const POS_BIT_D2 As Integer = 1
Public Const POS_BIT_D3 As Integer = 2
Public Const POS_BIT_D4 As Integer = 3
Public Const POS_BIT_C1 As Integer = 8
Public Const POS_BIT_C2 As Integer = 9
Public Const POS_BIT_C3 As Integer = 10
Public Const POS_BIT_C4 As Integer = 11

'--- nombre de mots d'une fiche redresseur dans l'API ---
Public Const NBR_MOTS_FICHE_REDRESSEUR_API As Integer = 10
                            
'--- positions des bits utilisés pour l'états des redresseurs ---
Public Const POS_BIT_FIN_GAMME As Integer = 0
Public Const POS_BIT_DEFAUT_REDRESSEUR As Integer = 7
Public Const POS_BIT_AUTO_REDRESSEUR As Integer = 8
Public Const POS_BIT_CATHODIQUE_REDRESSEUR As Integer = 9
Public Const POS_BIT_ANODIQUE_REDRESSEUR As Integer = 10
Public Const POS_BIT_MARCHE_REDRESSEUR As Integer = 11

'--- positions des bits utilisés pour les annnexes ---
Public Const POS_BIT_PV_VENTIL As Integer = 9
Public Const POS_BIT_GV_VENTIL As Integer = 10
Public Const POS_BIT_DISJONCTION_VENTIL As Integer = 15
                        
Public Const POS_BIT_VOLET_COMPENSATION_OUVERT As Integer = 0
Public Const POS_BIT_VOLET_COMPENSATION_FERME As Integer = 1
Public Const POS_BIT_DISJONCTION_VOLET_COMPENSATION As Integer = 2

Public Const POS_BIT_MARCHE_SURPRESSEUR_AIR As Integer = 0
Public Const POS_BIT_DISJONCTION_SURPRESSEUR_AIR As Integer = 1

Public Const POS_BIT_MARCHE_ROTATION_TONNEAU_CUVES As Integer = 0
Public Const POS_BIT_DISJONCTION_ROTATION_TONNEAU_CUVES As Integer = 1

Public Const POS_BIT_DISJONCTION_POMPE_ANODISATION As Integer = 0
Public Const POS_BIT_DISJONCTION_SURPRESSEUR_EAU As Integer = 1
Public Const POS_BIT_DISJONCTION_ROTATION_TONNEAU As Integer = 2

'--- cas général de marche et d'arrêt ---
Public Const ARRET  As Integer = 0
Public Const MARCHE  As Integer = 1

'--- cas général du manuel et de l'automatique ---
Public Const MANU  As Integer = 0
Public Const AUTO  As Integer = 1

'--- mode de décodage des chaines retournées par l'API ---
Public Const ENTIER  As Integer = 1
Public Const HEXADECIMAL  As Integer = 2
Public Const BINAIRE As Integer = 3
            
'--- code de l'exclusion d'un redresseur ---
Public Const CODE_EXCLUSION_REDRESSEUR As Integer = 14
            
'--- pour un cycle de 24 heures (journées types et programmateur cyclique) ---
Public Const NBR_TOPS_POSSIBLES As Integer = 50

'--- pour le fonctionnement du programmateur cyclique ---
Public Const NBR_JOURS_PROG_CYCLIQUE As Integer = 15

Public Const NUMZONE_ANO As Integer = 15

'--- pour les graphes de production ---
Public Const CANAL_DEPART_TRACABILITE As Integer = 200            'canal de départ pour la traçabilité
Public Const NBR_POINTS_MAXI_TRACABILITE As Long = 1900       'le nombre maxi de points possibles à tracer sur l'outil graphique est de 3800 points (2 courbes x 1900 points)
Public Const POURCENT_AVANT_TRACABILITE As Single = 0.03       'pourcentage de tolérance avant de mémoriser un point

'--- actions ---
Public Const NUM_ACTION_NOP As Integer = 0                                                    'NOP = Pas d'opération

Public Const NUM_ACTION_NIVEAU_BAS As Integer = 201                                  'NB = Niveau bas
Public Const NUM_ACTION_NIVEAU_INTERMEDIAIRE As Integer = 202              'NI = Niveau intermédiaire
Public Const NUM_ACTION_NIVEAU_HAUT As Integer = 215                                'NH = Niveau haut

Public Const NUM_ACTION_FCY As Integer = 8000                                              'FCY = Fin de cycle d'un pont

'--- émulateur des commandes ---
Public Const NBR_COMMANDES_EMULATEUR As Integer = 10         'nombre de commandes contenues
                                                                                                               'dans le fichier d'émulation

'*** CONSTANTES CHAINES ***

'--- programme ---
Public Const INDICATIF_PROGRAMME As String = " TECAL VERBRUGGE - ANODISATION - "
Public Const INDICATIF_ZONE_USINE As String = "ANODISATION"                   'indicatif de zone dans l'usine

'--- mot de passe système ---
Public Const MOT_DE_PASSE_SYSTEME As String = "CDB"  'permet l'accés sans connaitre le mot de passe opérateur

'--- connexions aux bases de données ---
Public Const CST_PARAMETRES_CONNEXION_BD_ANODISATION_TEST_SQL As String = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=ANODISATION;Uid=sa; Pwd=sa;Data Source=SRV2003\SQLEXPRESS;Connect Timeout=3;"
'Public Const CST_PARAMETRES_CONNEXION_BD_ANODISATION_SQL As String = "Provider=SQLNCLI11;Server=SRV-APP-ANOD\SQLEXPRESS;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"
Public Const CST_PARAMETRES_CONNEXION_BD_ANODISATION_SQL As String = "Provider=SQLNCLI11;Server=VB-LANLIGNE2-20\SQLEXPRESSANO;Database=ANODISATION;Uid=sa; Pwd=Jeff_nenette;"


Public Const CST_PARAMETRES_CONNEXION_BD_CLIPPER_HF As String = "Provider=PCSoft.HFSQL;Initial Catalog=TECAL-VERBRUGGE;User ID=admin;Data Source=VBVSE001:4924;"
Public Const CST_PARAMETRES_CONNEXION_BD_CLIPPER_TEST_HF As String = "Provider=PCSoft.HFSQL;Initial Catalog=TECAL-VERBRUGGE-TEST;User ID=admin;Data Source=VBVSE001:4924;"
Public Const CST_PARAMETRES_CONNEXION_BD_CLIPPER_TEST_ACCESS As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Anodisation\Base de données\CLIPPER.mdb;Persist Security Info=False;"

Public Const PARAMETRES_CONNEXION_BD_SAGE_SQL As String = ""

'--- fichiers ---
Public Const FIC_CONFIGURATION  As String = "config.ini"
Public Const FIC_REGULATION  As String = "Régulation.txt"
Public Const FIC_JOURNEES_TYPES  As String = "Journées types.txt"
Public Const FIC_PROG_CYCLIQUE  As String = "Programmateur cyclique.txt"
Public Const FIC_ANNEXES  As String = "Annexes.txt"
Public Const FIC_ETATS_POSTES  As String = "Etats des postes.txt"
Public Const FIC_PARAMETRES_LIGNE As String = "Paramètres de la ligne.txt"
Public Const FIC_SIGNALISATION_DEFAUTS As String = "Signalisation des défauts.txt"

Public Const FIC_BAINS_ANODISATION  As String = "Bains anodisation.txt"

'--- tables de la base de données de Sage ---
Public Const TABLE_POINTAGES_BAINS As String = "PointagesBains"

'--- tables de la base de données d'anodisation ---
Public Const TABLE_GAMMES_ANODISATION As String = "GammesAnodisation"
Public Const TABLE_DETAILS_GAMMES_Anodisation As String = "DetailsGammesAnodisation"
Public Const TABLE_DETAILS_CHARGES_PRODUCTION As String = "DetailsChargesProduction"
Public Const TABLE_DETAILS_GAMMES_ANODISATION_PRODUCTION As String = "DetailsGammesProduction"
Public Const TABLE_DETAILS_PHASES_PRODUCTION As String = "DetailsPhasesProduction"
Public Const TABLE_DETAILS_FICHES_PRODUCTION As String = "DetailsFichesProduction"
Public Const TABLE_TRACABILITE_ALARMES As String = "TraçabiliteAlarmes"

Public Const TABLE_IMP_DETAILS_CHARGE_1 As String = "ImpDetailsCharge1"
Public Const TABLE_IMP_DETAILS_DETAILS_CHARGE_1 As String = "ImpDetailsDetailsCharge1"
Public Const TABLE_IMP_DETAILS_REFERENCES_CLIENTS_1 As String = "ImpDetailsReferencesClients1"

Public Const TABLE_IMP_GAMMES_ANODISATION_PRODUCTION_1 As String = "ImpGammesProduction1"
Public Const TABLE_IMP_DETAILS_GAMMES_ANODISATION_PRODUCTION_1 As String = "ImpDetailsGammesProduction1"

Public Const TABLE_IMP_TRACABILITE_CHARGE_1 As String = "ImpTraçabiliteCharge1"
Public Const TABLE_IMP_DETAILS_TRACABILITE_CHARGE_1 As String = "ImpDetailsTraçabiliteCharge1"

Public Const TABLE_IMP_PRODUCTION_PAR_JOUR_1 As String = "ImpProductionParJour1"
Public Const TABLE_IMP_DETAILS_PRODUCTION_PAR_JOUR_1 As String = "ImpDetailsProductionParJour1"

Public Const TABLE_IMP_ALARMES_LIGNE_1 As String = "ImpAlarmesLigne1"
Public Const TABLE_IMP_DETAILS_ALARMES_LIGNE_1 As String = "ImpDetailsAlarmesLigne1"

'--- ATTENTION textes des codes pour le calcul des temps ---
Public Const CODE_NIVEAU_BAS As String = "NB"
Public Const CODE_NIVEAU_INTERMEDIAIRE As String = "NI"
Public Const CODE_NIVEAU_HAUT As String = "NH"

Public Const CODE_TEMPO As String = "TEMPO"
Public Const CODE_TEMPO_EGOUTTAGE As String = "TEMPO_EGOUT"
Public Const CODE_TEMPO_STABILISATION As String = "TEMPO_STAB"

Public Const CODE_OUVERTURE_COUVERCLES As String = "OCO"
Public Const CODE_CONTROLE_COUVERCLES_OUVERTS As String = "CCO"
Public Const CODE_FERMETURE_COUVERCLES As String = "FCO"

Public Const CODE_DESCENTE_ACCROCHES As String = "DEAC"
Public Const CODE_MONTEE_ACCROCHES As String = "MOAC"

Public Const CODE_ZONE_ANODISATION As String = "C13 à C16"

Public Const CODE_ARRET_AGITATION As String = "AAGIT"
Public Const CODE_MARCHE_AGITATION As String = "MAGIT"

Public Const CODE_ARRET_SECHOIR As String = "ASECHOIR"
Public Const CODE_MARCHE_SECHOIR As String = "MSECHOIR"

Public Const CODE_FIN_DE_CYCLE As String = "FCY"

'--- textes prédéfinis ---
Public Const PAS_DE_TEMPS As String = "-"
Public Const TEXTE_ANODIQUE As String = "ANODIQUE"
Public Const TEXTE_CATHODIQUE As String = "CATHODIQUE"

'--- séparateurs pour le cryptage des données en une seule ligne ---
Public Const SEPARATEUR_POSTES As String * 1 = ","
Public Const SEPARATEUR_PREMISSES As String * 1 = "-"
Public Const SEPARATEUR_NUM_DEFAUTS As String = "-"

'--- formats standards de nombre ---
Public Const FORMAT_NATURELS_2_CHIFFRES As String = "#0"
Public Const FORMAT_NATURELS_3_CHIFFRES As String = "##0"
Public Const FORMAT_NATURELS_4_CHIFFRES As String = "###0"

'--- formats divers pour l'affichage de variables ---
Public Const FORMAT_NUM_GAMME_ANODISATION As String * 6 = "000000"
Public Const FORMAT_NUM_FICHE_PRODUCTION As String * 8 = "00000000"
Public Const FORMAT_NUM_CDE_INTERNE As String * 8 = ""

Public Const FORMAT_COMPENSATION As String * 6 = "###0"

Public Const FORMAT_DELAI_SUP_STABILISATION_CHARGE As String * 2 = "#0"

Public Const FORMAT_DATE_HEURE_1 As String = "dd/mm/yyyy à hh:nn:ss"

Public Const FORMAT_TEMPERATURE_1_DECIMALE As String = "##0.0"
Public Const FORMAT_TEMPERATURE_1_DECIMALE_UNITE As String = "##0.0 °C"
Public Const FORMAT_TEMPERATURE_COMPACTE_1_DECIMALE_UNITE As String = "##0.0°C"

Public Const FORMAT_INTENSITE_ENTIER As String = "###0"
Public Const FORMAT_INTENSITE_1_DECIMALE As String = "###0.0"
Public Const FORMAT_INTENSITE_ENTIER_UNITE As String = "###0 A"                  'format avec l'unité de mesure
Public Const FORMAT_INTENSITE_1_DECIMALE_UNITE As String = "###0.0 A"      'format avec l'unité de mesure

Public Const FORMAT_TENSION_1_DECIMALE As String = "#0.0"
Public Const FORMAT_TENSION_2_DECIMALES As String = "#0.00"
Public Const FORMAT_TENSION_1_DECIMALE_UNITE As String = "#0.0 V"              'format avec l'unité de mesure
Public Const FORMAT_TENSION_2_DECIMALES_UNITE As String = "#0.00 V"          'format avec l'unité de mesure

Public Const FORMAT_ANALYSEUR As String = "##0.00"
Public Const FORMAT_ANALYSEUR_UNITE As String = "##0.00 g/l"   'format avec l'unité de mesure

Public Const FORMAT_POIDS_SOULEVE As String = "##0.0 kg"





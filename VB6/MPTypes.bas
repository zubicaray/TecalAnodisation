Attribute VB_Name = "MPTypes"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES TYPES PUBLIQUES
' Nom                    : MPTypes.bas
' Date de création : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z


Type VbHeader
    szVbMagic               As String * 4
    wRuntimeBuild           As Integer
    szLangDll               As String * 14
    szSecLangDll            As String * 14
    wRuntimeRevision        As Integer
    dwLCID                  As Long
    dwSecLCID               As Long
    lpSubMain               As Long
    lpProjectInfo           As Long
    fMdlIntCtls             As Long
    fMdlIntCtls2            As Long
    dwThreadFlags           As Long
    dwThreadCount           As Long
    wFormCount              As Integer
    wExternalCount          As Integer
    dwThunkCount            As Long
    lpGuiTable              As Long
    lpExternalCompTable     As Long
    lpComRegisterData       As Long
    bszProjectDescription   As Long
    bszProjectExeName       As Long
    bszProjectHelpFile      As Long
    bszProjectName          As Long
End Type
'***************************************************************************************************************************
'                                                                                   ACTIONS
'***************************************************************************************************************************

'--- enregistrement du type des actions ---
Public Type EnrActions
    NumAction As Long                          'N° de l'action
    CodeAction As String                        'Code de l'action
    LibelleAction As String                     'libellé de l'action
    ParametreOuiNon As Boolean         'Indique si l'action a un paramètre
    LibelleParametre As String              'libellé du paramètre
End Type

'***************************************************************************************************************************
'                                                                          CYCLE D'UN PONT
'***************************************************************************************************************************

'--- cycle d'un pont ---
Public Type VarCyclePont
    TActions(0 To NBR_LIGNES_CYCLES_PONTS) As EnrActions  'tableau contenant le cycle du pont
    PtrActions As Integer                                                                   'pointeur des actions du pont
    NumAction As Long                                                                      'numéro de l'action en cours du pont
End Type

'***************************************************************************************************************************
'                                                                   TEMPS DE MOUVEMENTS
'***************************************************************************************************************************

'--- variable d'analyse des mouvements ---
Public Type VarAnalyseMouvements
    EtatMouvement As ETATS_MOUVEMENTS          '0 = Pas de mouvement (état de repos avant le début du mouvement)
                                                                                  '1 = Le mouvement est en cours
                                                                                  '2 = Fin du mouvement
                                                                                  '3 = Calcul du temps réel du mouvement puis remise à 0
    DateDebutMouvement As Date                            'date de début d'un mouvement
    DateFinMouvement As Date                                 'date de fin d'un mouvement
End Type

'***************************************************************************************************************************
'                                                                               PREMISSES
'***************************************************************************************************************************

'--- variable du type des prémisses pour la création du tableau des prémisses ---
Public Type VarPremisses
    NumPont As Integer                                              'n° du pont concerné défini comme règle au départ
    NumPontIA As Integer                                           'n° du pont choisi par le moteur d'inférence
    PremisseCodee As String                                     'prémisse codée
    PremisseDecodee As String                                 'prémisse décodée
    TempsCycleSecondes As Long                            'temps du cycle en secondes
End Type

'--- variable du type des prémisses pour le calcul du temps de cycle ---
Public Type VarPremissesTempsCycle
    NumAction As Integer                      'n° de l'action
    CodeAction As String                       'code de l'action
    ParametreOuiNon As Boolean        'paramètre oui ou non
    Parametre As String                        'paramètre en fonction de l'action
    LibelleAction As String                    'libellé de l'action
End Type

'***************************************************************************************************************************
'                                                                             REDRESSEURS
'***************************************************************************************************************************

'--- détails des phases U et I pour renseigner le redresseur en entrant dans le bain ---
Public Type DetailsPhases
    TempsPhase As Integer                                                            'temps de la phase
    UPhase As Single                                                                     'tension de la phase
    IPhase As Single                                                                       'intensité de la phase
End Type

'--- enregistrement du type redresseurs ---
Public Type EnrRedresseurs
    NumRedresseur As Integer                                                       'N° du redresseur
    NomRedresseur As String                                                         'nom du redresseur
    LibelleRedresseur As String                                                      'libellé du redresseur
    UMaxiRedresseur As Single                                                      'tension maximale donnée par le constructeur du redresseur
    IMaxiRedresseur As Single                                                       'courant maximum donné par le constructeur du redresseur
End Type

'--- entrées des redresseurs ---
Public Type EntreesRedresseurs

    M_DefautGeneral As Boolean                                                    'défaut général
    
    M_DelaiTropLongMarcheRedresseur As Boolean                     'délai trop long de mise en marche du redresseur (contrôle sur l'intensité)

    M_IntensiteNonAtteinte As Boolean                                           'intensité demandée non atteinte
    
    M_IntensiteInstable As Boolean                                                 'intensité instable en cours de fonctionnement sur la phase 4
    
End Type

'--- numéros des défauts ---
Public Type NumDefautsRedresseurs
    
    NumDefautDefautGeneral As Integer                                         'numéro de défaut du défaut général
    
    NumDefautDelaiTropLongMarcheRedresseur As Integer          'numéro de défaut du délai trop long de mise en marche du redresseur (contrôle sur l'intensité)
   
    NumDefautIntensiteNonAtteinte As Integer                                'numéro de défaut du intensité demandée non atteinte
   
    NumDefautIntensiteInstable As Integer                                      'numéro de défaut du intensité instable en cours de fonctionnement sur la phase 4
   
End Type

'--- états des redresseurs ---
Public Type EtatsRedresseurs
    
    TNumDefauts As NumDefautsRedresseurs                                'numéros des défauts
    
    TEntreesAPI As EntreesRedresseurs                                          'entrées des redresseurs (automate)
    
    DefinitionRedresseur As EnrRedresseurs                                  'définition d'un redresseur
    
    ModeRedresseur As MODES_REDRESSEUR                             'mode du redresseur
    EtatRedresseur As ETATS_REDRESSEUR                                  'état électrique du redresseur
    SensRedresseur As SENS_REDRESSEUR                                 'sens du redresseur
    
    EtatsMarcheArret As String                                                          'états de marche et d'arrêt
    Etats1 As String                                                                           'états 1 du redresseur
    Etats2 As String                                                                           'états 2 du redresseur
    
    DemandesDuPC As Integer                                                         'demande du PC
    RetoursVersPC As Integer                                                           'retours vers PC (les valeurs retournées doivent correspondre aux valeurs demandées)
    IDemandePC As Integer                                                               'Intensité demandé par le PC
    
    NumCharge As Integer                                                                 'numéro de la charge traité par le redresseur
    
    NumPhaseEnCours As PHASES_GAMMES_REDRESSEURS      'numéro de la phase en cours (gamme d'anodisation en automatique)
    TempsPhaseEnCours As Integer                                                  'temps de la phase en cours
    TempsEcoulePhaseEnCours As Integer                                       'temps écoulé de la phase en cours
    
    DebutCycle As Boolean                                                                'TRUE = Indique le début d'un cycle, FALSE = Pas de cycle en cours
    ControleFinCycle As Boolean                                                       'TRUE = Indique la fin d'un cycle, FALSE = Cycle en cours
    
    TempsAjouteSurIFaible As Integer                                                'temps ajouté sur une intensité plus faible que prévue (panne de redresseur)
    TempsTotalCycle As Integer                                                         'temps total du cycle en secondes lu dans l'automate
    TempsRestantCycle As Integer                                                     'temps restant du cycle en secondes lu dans l'automate
    TempsTotalise As Integer                                                             'temps totalisé
    
    ModeUouI As MODES_U_OU_I                                                     'mode de travail du redresseur U(tension)=0, I(intensité)=1
    
    TDetailsPhases(PHASES_GAMMES_REDRESSEURS.PH_T1 To _
                              PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases    'détails des phases
    
    U As Single                                                                                   'tension mesurée en volts
    I As Integer                                                                                    'courant mesurée en ampères
    
    ConsigneU As Single                                                                    'consigne en tension
    ConsigneI As Integer                                                                    'consigne en intensité

    Ah As Single                                                                                  'Ah calculée
    
End Type

'***************************************************************************************************************************
'                                                                          ETATS DE LA LIGNE
'***************************************************************************************************************************

'--- numéros des défauts ---
Public Type NumDefautsLigne
    
    NumDefautArretGeneral As Integer
    
    NumDefautArretUrgence As Integer
    
    NumDefautPortillonsLigneVie As Integer
    
    NumDefautStopLigne As Integer
    
    NumDefautArretUrgenceP1 As Integer
    
    NumDefautSecuriteP1 As Integer
    
    NumDefautArretUrgenceP2 As Integer
    
    NumDefautSecuriteP2 As Integer
    
    NumDefautManqueTension As Integer
    
    NumDefautManqueAir As Integer
    
End Type

Public Type EtatsLigne
    
    TNumDefauts As NumDefautsLigne  'numéros des défauts
    
    MarcheGenerale As Boolean          'marche générale de la ligne
    ArretGeneral As Boolean                'arrêt général de la ligne
    ArretUrgence As Boolean               'arrêt d'urgence
    PortillonsLigneVie As Boolean       'portillons et ligne de vie
    StopLigne As Boolean                    'stop ligne
    
    ArretUrgenceP1 As Boolean           'arrêt d'urgence du pont 1
    ArretUrgenceP2 As Boolean           'arrêt d'urgence du pont 2
    
    SecuriteP1 As Boolean                   'sécurité du pont 1
    SecuriteP2 As Boolean                   'sécurité du pont 2
    
    ManqueTension As Boolean           'manque de tension
    ManqueAir As Boolean                    'manque d'air
    AcquittementsDefauts As Boolean  'pour l'acquittement des défauts
    FrontMontantDefauts As Boolean    'front montant des défauts

End Type

'***************************************************************************************************************************
'                                                                                   PONTS
'***************************************************************************************************************************

'--- paramètres des cycles des ponts ---
Public Type ParametresCyclesPonts
    NumPosteDepart As Integer                                    'poste de départ d'un cycle d'un pont
    NumPosteArrivee As Integer                                   'poste d'arrivée d'un cycle d'un pont
    TypeCycle As TYPES_CYCLES                               'déplacement ou transfert de charge
    DelaiSupStabilisationChargeSecondes As Integer 'délai supplémentaire de stabilisation de la charge
    TempsEgouttageSecondes As Integer                    'temps d'égouttage en secondes
End Type

'--- cycles des ponts ---
Public Type CyclesPonts
    NumAction As Integer                   'N° de l'action
    Parametre As String                     'Paramètre en fonction de l'action
    EtatParametre As String               'Etat des paramètres (indique le temps qu'il reste ou autre)
End Type

'--- variable du type pointeur de l'action et action en cours données par l'API ---
Public Type VarPtrEtActionEnCoursAPI
    PtrAction As Integer                      'pointeur de l'action
    NumAction As Integer                   'N° de l'action
    Parametre As Integer                   'paramètre
End Type

'--- entrées des ponts API ---
Public Type EntreesPontsAPI
    
    M_MoteurTourneTrlPont As Boolean                                      'le moteur tourne de la translation gauche du pont

    M_MoteurTourneLevPont As Boolean                                    'le moteur tourne du levage du pont
    
    M_MarquageAxeTrL As Boolean                                            'marquage axe de la translation
    M_MarquagePVTrL As Boolean                                              'marquage de la petite vitesse de la translation
    M_MarquageMVTrL As Boolean                                             'marquage de la moyenne vitesse de la translation
    M_MarquageArretTrL As Boolean                                          'marquage d'arrêt de la translation
    
    M_MarquageAxeLev As Boolean                                           'marquage axe du levage
    
    M_MemDemandeIsocentrage As Boolean                             'mémoire de demande d'isocentrage aprés un glissement important
    
    M_DefautPresencePicece  As Boolean                                  'indique poste occupé sur Cellule
    
    M_MarquageAxeLevPont As Boolean                                    'marquage axe du levage du pont
    M_ErreurPointeur As Boolean                                                'erreur sur le pointeur des actions ou sur le code des actions
    
    M_AccrochesEnHaut As Boolean
    M_AccrochesEnBas As Boolean
    
    E_NiveauHaut As Boolean
    E_NiveauIntermediaire As Boolean
    E_NiveauBas As Boolean
    
    M_DefautVariateurTrlPont As Boolean                                   'défaut variateur de la translation gauche du pont
    M_AxeNonReferenceTrlPont As Boolean                               'axe non référencé de la translation gauche du pont
    M_SurcourseTrlAvant As Boolean                                          'surcourse de la translation avant
    M_SurcourseTrlArriere As Boolean                                        'surcourse de la translation arrière
    M_DefautVariateurLevPont As Boolean                                 'défaut variateur du levage du pont
    M_AxeNonReferenceLevPont As Boolean                              'axe non référencé du levage du pont
    M_SurcourseLevBas As Boolean                                           'surcourse levage bas
    M_SurcourseLevHaut As Boolean                                          'surcourse levage haut
    
    M_DelaiTropLongDescenteAccroches As Boolean
    M_DelaiTropLongMonteeAccroches As Boolean
    
End Type

'--- sorties des ponts API ---
Public Type SortiesPontsAPI
    
    A_AntiCollision As Boolean                                                    'sortie API du surcourse avant pour le pont 1 et arrière du pont 2
    
    S_EVMonteeAccroches As Boolean                                       'électro-vanne de montée des accroches (libération de la charge)
    S_EVDescenteAccroches As Boolean                                    'électro-vanne de descente des accroches (prise de la charge)

End Type

'--- temps des mouvements ---
'tous les temps sont en secondes
Public Type TempsMouvementsPonts
    
    TTempsTranslation(POSTES.P_CHGT_1 To DERNIER_POSTE, POSTES.P_CHGT_1 To DERNIER_POSTE) As Single
                                                                                          'temps de déplacement du pont d'un poste de départ vers
                                                                                          'un poste d'arrivée
    
    TempsAccrochesChargeVersHaut As Single               'temps accroches charges vers le haut
    TempsAccrochesChargeVersBas As Single                'temps accroches charges vers le bas
    
    TempsDescenteHautVersBas As Single                      'temps en descente du niveau haut vers le niveau bas
    TempsDescenteIntermediaireVersBas As Single        'temps en descente du niveau intermédiaire vers le niveau bas

    TempsMonteeBasVersIntermediaire As Single           'temps en montée du niveau bas vers le niveau intermédiaire
    TempsMonteeBasVersHaut As Single                         'temps en montée du niveau bas vers le niveau haut

End Type

'--- numéros des défauts ---
Public Type NumDefautsPonts
    
    NumDefautDefautVariateurTrLPont As Integer
    
    NumDefautAxeNonReferenceTrlPont As Integer

    NumDefautSurcourseTrlAvant As Integer
    NumDefautSurcourseTrlArriere As Integer
    
    NumDefautDefautVariateurLevPont As Integer
    NumDefautAxeNonReferenceLevPont As Integer

    NumDefautSurcourseLevBas As Integer
    NumDefautSurcourseLevHaut As Integer

    NumDefautDelaiTropLongDescenteAccroches As Integer
    NumDefautDelaiTropLongMonteeAccroches As Integer
    
    NumDefautAntiCollision As Integer
    NumDefautFinDeZone As Integer
    
    NumDefautPresencePiece As Integer
    NumDefautDefautLaser As Integer
    
End Type

'--- états des ponts ---
Public Type EtatsPonts
    
    ModePont As MODES_PONTS                 'mode du pont de maintenance à automatique
    TypeSequence As TYPES_SEQUENCES 'type de séquence du pont (inconnu, cyclique, aléatoire)
    ControleParOperateur As Boolean           'FALSE=contrôle par IA, TRUE=contrôle par opérateur (uniquement en mode automatique)
    
    TypesAffichagesCyclesPonts As Boolean    'pour l'affichage des cycles de ponts
                                                                           'FALSE = Cycle actuel, TRUE = Prochain cycle

    TParametresCyclesPonts(CYCLES.C_ACTUEL To CYCLES.C_PROCHAIN) As ParametresCyclesPonts
                                                                           'paramètres des cycles des ponts (poste de départ et arrivée, etc ...)
    TCyclesPonts(CYCLES.C_ACTUEL To CYCLES.C_PROCHAIN, 1 To NBR_LIGNES_CYCLES_PONTS) As CyclesPonts
                                                                     'cycles des ponts

    PtrEtActionEnCoursAPI As VarPtrEtActionEnCoursAPI
                                                                     'pointeur de l'action et action en cours retournée par l'automate
                                                                     '(ne concerne que le cycle actuel)
    
    TEntreesAPI As EntreesPontsAPI             'entrées des ponts (automate)
    TSortiesAPI As SortiesPontsAPI               'sorties des ponts (automate)
    
    TTempsMouvements As TempsMouvementsPonts    'pour la mémorisation des temps de mouvements
    
    TNumDefauts As NumDefautsPonts          'numéros de défauts
    UnDefautAuMoinsSignale As Boolean      'un défaut au moins est signalé
    
    PositionActuelleLaserTrlPont As Long      'position actuel du laser de la translation gauche
    PositionCibleLaserTrlPont As Long           'position cible du laser de la translation gauche
    
    PositionActuelleCodeurTrlPont As Long    'position actuelle du codeur de la translation gauche
    PositionCibleCodeurTrlPont As Long        'position cible du codeur de la translation gauche
    
    PositionActuelleCodeurLevPont As Long  'position actuelle du codeur de levage
    PositionCibleCodeurLevPont As Long      'position cible du codeur de levage
     
    PosteActuel As POSTES                           'numéro du poste actuel
    PosteDestination As POSTES                  'numéro du poste de destination
    SensX As SENS_X                                    '1=sens avant, -1=sens arrière, 0=arrêt au poste

    NiveauActuel As NIVEAUX_PONTS           'numéro du niveau actuel (de 201 à 215)
    NiveauDestination As NIVEAUX_PONTS  'numéro du niveau de destination
    SensY As SENS_Y                                    '1=sens montée, -1=sens descente, 0=arrêt au niveau

    EtatsAccrochesCharge As ETATS_ACCROCHES    '0 = Accroches de la charge en haut, 1 = Accroches de la charge en bas

    PoidsSouleve As Single                            'poids soulevé

    Condamnation As Boolean                        'TRUE=Pont condamné
    NumCharge As Integer                              'N° de la charge sur le pont
    OptionsGamme1 As Integer                      'options de la gamme partie 1
    OptionsGamme2 As Integer                      'options de la gamme partie 2
    
    Alarmes As String                                     'N° des alarmes du pont (disjonction ou séquence non effectuée)

End Type

'***************************************************************************************************************************
'                                                                                   POSTES
'***************************************************************************************************************************

'--- enregistrement du type postes ---
Public Type EnrPostes
    
    NumPoste As Integer                                    'numéro du poste dans la ligne
    NomPoste As String                                      'nom du poste dans la ligne
    LibellePoste As String                                  'libellé complet du poste dans la ligne
    
    AvecTemps As Boolean                               'FALSE=Pas de temps au poste (chargement, déchargement ...)
                                                                         'TRUE=Avec un temps obligatoire (cas des bains)
    
    RespectTempsObligatoire As Boolean        'FALSE=Le temps de bain peut admettre d'être dépassé
                                                                         'TRUE=Le temps de bain doit être respecté
    
    AvecEgouttage As Boolean                          'FALSE=Pas d'égouttage au poste, TRUE=Avec égouttage au poste
    
    PresenceCouvercles As Boolean                 'FALSE=Pas de couvercles, TRUE=Présence de couvercles
    PresenceRedresseur As Boolean                'FALSE=Pas de redresseur, TRUE=Présence d'un redresseur
    PresenceAgitationBain As Boolean              'FALSE=Pas d'agitation du bain, TRUE=Présence d'une agitation du bain
    
    XAxePosteLigne As Long                              'X de l'axe de poste dans la ligne (valeur du lecteur laser)
    
    XAxePosteSynoptique As Long                     'X de l'axe de poste dans le synoptique
    
    XInferieurPosteSynoptique As Long             'X inférieur du rectangle limitant le poste dans le synoptique
    YInferieurPosteSynoptique As Long             'Y inférieur du rectangle limitant le poste dans le synoptique
    XSuperieurPosteSynoptique As Long           'X supérieur du rectangle limitant le poste dans le synoptique
    YSuperieurPosteSynoptique As Long           'X supérieur du rectangle limitant le poste dans le synoptique
    
    XInferieurLibellePosteSynoptique As Long   'X inférieur du rectangle limitant le libellé du poste dans le synoptique
    YInferieurLibellePosteSynoptique As Long   'Y inférieur du rectangle limitant le libellé du poste dans le synoptique
    XSuperieurLibellePosteSynoptique As Long 'X supérieur du rectangle limitant le libellé du poste dans le synoptique
    YSuperieurLibellePosteSynoptique As Long 'X supérieur du rectangle limitant le libellé du poste dans le synoptique

End Type

'--- états des postes ---
Public Type EtatsPostes
    
    DefinitionPoste As EnrPostes                                       'définition du poste
    
    Condamnation As Boolean                                            'TRUE=Poste condamné
                                                                                           '           pas de prise ni de dépose de charge autorisées
                                                                                           'FALSE=Poste en fonctionnement normal
                                                                                           '           prise et dépose de charge autorisées
    
    NumCharge As Integer                                                  'N° de la charge dans le poste
    
    EtatsChariots As ETATS_CHARIOTS                             'états des chariots
    
    Alarmes As String                                                          'N° des alarmes du poste (disjonction, etc...)
    
End Type

'***************************************************************************************************************************
'                                                               GESTION DES TEMPERATURES
'***************************************************************************************************************************

'--- variable de gestion des températures ---
Public Type VarTemperatures
    TempActuelle As Single                     'température actuelle de la cuve (valeur retournée par l'automate)
    TempVeille As Single                         'température de veille (opérateur)
    TempProduction As Single                 'température de production normale (opérateur)
    EcartInferieurRegul As Single            'écart inférieur de régulation (opérateur)
    EcartSuperieurRegul As Single          'écart supérieur de régulation (opérateur)
    EcartInferieurAlarme As Single          'écart inférieur d'alarme (opérateur)
    EcartSuperieurAlarme As Single        'écart supérieur d'alarme (opérateur)
End Type

'***************************************************************************************************************************
'                                                                                    CUVES
'***************************************************************************************************************************

'--- enregistrement du type cuves ---
Public Type EnrCuves
    
    NumCuve As Integer                                                         'numéro de la cuve dans la ligne
    NomCuve As String                                                           'nom de la cuve dans la ligne
    LibelleCuve As String                                                        'libellé complet de la cuve dans la ligne
   
    GestionAPI As Boolean                                                      'indique si la cuve est gérée par l'automate
    
    PresencePompe As Boolean                                             'FALSE=Pas de pompe, TRUE=Présence d'une pompe
    NbrChauffages As Integer                                                  'nombre de chauffages (0=cuve non chauffée)
    PresenceRefroidissementBain As Boolean                      'FALSE=Pas de refroidissement du bain, TRUE=Refroidissement du bain
    PresenceNiveauBas As Boolean                                      'FALSE=Pas de niveau bas, TRUE=Présence d'un niveau bas
    PresenceNiveauHaut As Boolean                                     'FALSE=Pas de niveau haut, TRUE=Présence d'un niveau haut
    PresenceEVEau As Boolean                                              'FALSE=Pas d'électro-vanne d'eau, TRUE=Présence d'une électro-vanne d'eau
    PresenceAnalyseurAnodisation As Boolean                     'FALSE=Pas d'analyseur d'anodisation, TRUE=Présence d'un analyseur d'anodisation

End Type

'--- caractéristiques générales de toutes les cuves ---
Public Type CaracteristiquesCuves
    
    DefinitionCuve As EnrCuves                 'définition d'une cuve

End Type

'--- entrées des cuves API ---
Public Type EntreesCuvesAPI
    
    E_ManuAutoRegulation As Boolean
    
    E_CouverclesOuverts As Boolean
    E_CouverclesFermes As Boolean
    
    E_NiveauTresBas As Boolean
    E_NiveauIntermediaireBas As Boolean
    E_NiveauIntermediaireHaut As Boolean
    E_NiveauTresHaut As Boolean
    
    E_DefautChauffage As Boolean
    E_DefautRefroidissement As Boolean
    E_DefautPompe As Boolean
    E_DefautEVEau As Boolean
    E_DelaiTropLongEVEau As Boolean
    E_DefautCouvercles As Boolean
    E_DefautAgitationBain As Boolean

    TemperatureTropBasse As Boolean
    TemperatureTropHaute As Boolean
    DefautPT100 As Boolean
    
    E_Analogique_Analyseur As Single                            'valeur analogique de l'analyseur

End Type

'--- sorties des cuves API ---
Public Type SortiesCuvesAPI
    
    S_Chauffage As Boolean                                             'sorties chauffage
    S_Dem_Chauffage As Boolean
    
    S_Refroidissement As Boolean
    S_Dem_Refroidissement As Boolean
    
    S_Pompe As Boolean
    S_Dem_Pompe As Boolean
    
    S_EVEau As Boolean
    S_Dem_EVEau As Boolean
    
    S_EVOuvertureCouvercles As Boolean
    S_Dem_EVOuvertureCouvercles As Boolean
    
    S_EVFermetureCouvercles As Boolean
    S_Dem_EVFermetureCouvercles As Boolean
    
    S_AgitationBain As Boolean
    S_Dem_AgitationBain As Boolean

End Type

'--- temps des mouvements ---
'tous les temps sont en secondes
Public Type TempsMouvementsCuves
    
    TempsOuvertureCouvercles As Single                        'temps d'ouverture des couvercles
    TempsFermetureCouvercles As Single                        'temps de fermeture des couvercles
    
End Type

'--- numéros des défauts ---
Public Type NumDefautsCuves
    
    NumDefautNiveauTresBas As Integer
    NumDefautNiveauTresHaut As Integer
    
    NumDefautDefautChauffage As Integer
    
    NumDefautDefautRefroidissement As Integer
    
    NumDefautDefautPompe As Integer
    
    NumDefautDefautEVEau As Integer
    
    NumDefautDelaiTropLongEVEau As Integer
    
    NumDefautDefautCouvercles As Integer
    
    NumDefautDelaiTropLongOuvertureCouvercles As Integer
    
    NumDefautDelaiTropLongFermetureCouvercles As Integer
    
    NumDefautDefautAgitationBain As Integer
    
    NumDefautTemperatureTropBasse As Integer
    NumDefautTemperatureTropHaute As Integer
    NumDefautDefautPT100 As Integer

End Type

'--- états des cuves gérées par l'automate ---
Public Type EtatsCuves
    
    DefinitionCuve As EnrCuves                                                                  'définition d'une cuve
    IndexAutomate As Integer
    API_Changements  As Boolean                                                             'TRUE=Indique un changement (hors états)
    API_Etats_1 As String * 16                                                                     'retour en binaire du mot de l'API contenant les états 1
    API_Etats_2 As String * 16                                                                     'retour en binaire du mot de l'API contenant les états 2
    
    TEntreesAPI As EntreesCuvesAPI                                                          'entrées des cuves (automate)
    TSortiesAPI As SortiesCuvesAPI                                                            'sorties des cuves (automate)
    
    TTempsMouvements As TempsMouvementsCuves                             'pour la mémorisation des temps de mouvements
    
    TNumDefauts As NumDefautsCuves                                                     'numéros de défauts
    UnDefautAuMoinsSignale As Boolean                                                   'un défaut au moins est signalé
    ListeNumDefautsSiCharge As String                                                     'liste des numéros de défauts en présence d'une charge
                                                                                                                   'si pas de charge alors cette variable est vide
     
    API_ModeProduction As MODES_PRODUCTION                                   'retour API du mode de production
    ModeProduction As MODES_PRODUCTION                                          'déterminer par le programmateur cyclique
    
    ModeRegulation As MODES_REGULATION                                          'mode de la régulation
    
    EtatsChauffage As ETATS_CHAUFFAGES                                              'états d'un chauffage
    EtatsRefroidissementBain As ETATS_REFROIDISSEMENT_BAIN        'états du froidissement d'un bain
    
    API_ModePompe As MODES_POMPES                                                 'retour API du mode de la pompe
    API_CyclePompe As CYCLES_POMPES                                                'retour API du cycle de la pompe
    ModePompe As MODES_POMPES                                                        'mode de la pompe déterminer par l'opérateur
    CyclePompe As CYCLES_POMPES                                                       'cycle de la pompe déterminer par le programmateur cyclique
    EtatsPompe As ETATS_POMPES                                                           'états de la pompe

    EtatsNiveaux As ETATS_NIVEAUX                                                         'états des niveaux
    
    EtatsEVEau As ETATS_EV_EAU                                                             'états de l'électro-vanne d'arrivée d'eau
    
    API_ModeCouvercles As MODES_COUVERCLES                                  'retour API du mode des couvercles
    API_CycleCouvercles As CYCLES_COUVERCLES                                 'retour API du cycle des couvercles (cycle des ponts)
    ModeCouvercles As MODES_COUVERCLES                                         'mode des couvercles déterminer par l'opérateur
    'pour les états des couvercles voir les postes
    
    Temperatures As VarTemperatures                                                       'valeurs des températures

End Type

'***************************************************************************************************************************
'                                                                        ZONES DE LA  LIGNE
'***************************************************************************************************************************

'--- enregistrement du type des zones de la ligne ---
Public Type EnrZones
    NumZone As Integer                     'N° de la zone
    Codezone As String                      'Code de la zone
    LibelleZone As String                   'Libellé de la zone
    NumPremierPoste As Integer       'N° du premier poste
    NomPremierPoste As String         'Nom du premier poste
    NumDernierPoste As Integer        'N° du dernier poste
    NomDernierPoste As String          'Nom du dernier poste
    NbrPostes As Integer                    'Nombre de postes concernés par la zone
End Type

'--- enregistrement du type des barres de la ligne ---
Public Type EnrBarres
    NumBarre As Integer                     'N° de la barre
    Libelle As String                   'Libellé de la barre
End Type

'***************************************************************************************************************************
'                                                                      GAMMES D'ANODISATION
'***************************************************************************************************************************

'--- enregistrement du type des détails des gammes d'anodisation ---
Public Type EnrDetailsGammesAnodisation
    
    NumLigne As Integer                                                  'n° de ligne
    NumZone As Integer                                                   'n° de zone
    
    TempsAuPosteSecondes As Long                              'temps au poste en secondes
    TempsAuPosteTexte As String                                    'temps au poste en texte au format HH:MM:SS
    
    TempsAlerteSecondes As Long                                   'temps d'alerte en secondes
    TempsAlerteTexte As String                                        'temps d'alerte en texte au format HH:MM:SS
    
    TempsEgouttageSecondes As Integer                        'temps d'égouttage en secondes
    TempsEgouttageTexte As String                                 'temps d'égouttage en texte au format MM:SS

    '********** UTILISER UNIQUEMENT EN PRODUCTION **********
    NumPosteReel As Integer                                           'n° de poste réel utilisé dans la zone
                                                                                         '(cas des postes multiples)
    
    DecompteDuTempsAuPosteReelSecondes As String 'représente la différence entre le temps théorique au poste
                                                                                         'et le temps réel passé dans le poste
                                                                                         'un nombre négatif apparait si la charge est resté plus
                                                                                         'longtemps dans le poste que le temps théorique prévu
                                                                                         'ATTENTION variable du type String volontairement
                                                                                         'Si "" alors il n'y a pas eu de temps de décompter
    
    DecompteDuTempsAlerteReelSecondes As String     'représente la différence entre le temps théorique d'alerte
                                                                                         'et le temps réel passé avant l'alerte
    
    FinDuTempsPosteReel As Boolean                            'TRUE = Indique la fin du temps au poste réel
    DebutAlertePosteReel As Boolean                              'TRUE = Indique le début de l'alertes au poste réel

End Type

'--- enregistrement du type 'matieres' ---
Public Type EnrMatieres
    Matiere As String                                                'matière
    TypeMatiere As String                                        'type de la matière
    CompositionMatiere As String                            'composition de la matière
    OrdrePourAffichage As Integer                           'ordre pour affichage
End Type

'--- enregistrement du type 'GammesAnodisation' (comprend également les détails) ---
Public Type EnrGammesAnodisation
    
    NumGamme As String                                              'n° de gamme
    DateCreationGamme As Date                                  'date de création de la gamme
    RefGamme As String                                                'référence de la gamme
    NomGamme As String                                              'nom de la gamme
    
    Designation As String                                               'désignation de la gamme d'anodisation
    
    TMatieresGamme(1 To NBR_MATIERES_MAXI_PAR_GAMME) As String    'tableau contenant les matières de la gamme
    
    TempsAvantPostePrincipalTexte As String             'temps avant Anodisation en texte au format HH:MM:SS
    TempsPostePrincipalTexte As String                      'temps au poste d'anodisation en texte au format HH:MM:SS
    TempsApresPostePrincipalTexte As String             'temps aprés Anodisation en texte au format HH:MM:SS
    TempsTotalPostesTexte As String                           'temps total des postes en texte au format HH:MM:SS
    TempsTotalEgouttagesTexte As String                    'temps total des égouttages en texte au format HH:MM:SS
    TempsTotalGammeTexte As String                          'temps total de la gamme en texte au format HH:MM:SS
    
    TempsAvantPostePrincipalSecondes As Long        'temps avant Anodisation en secondes
    TempsPostePrincipalSecondes As Long                 'temps au poste d'anodisation en secondes
    TempsApresPostePrincipalSecondes As Long        'temps aprés Anodisation en secondes
    TempsTotalPostesSecondes As Long                     'temps total des postes en secondes
    TempsTotalEgouttagesSecondes As Long              'temps total des égouttages en secondes
    TempsTotalGammeSecondes As Long                    'temps total de la gamme en secondes
    
    '************************* PASSAGE DANS LES BAINS **************************

    PassageAnodisation As Boolean                           'indique un passage dans un des bains d'anodisation
    PassageSpectro As Boolean                                  'indique un passage dans le bain de spectrocoloration
    PassageOr As Boolean                                           'indique un passage dans le bain d'or
    PassageNoir As Boolean                                        'indique un passage dans le bain de noir
    
    '*************************** GAMME REDRESSEUR *****************************
    ModeUouI As MODES_U_OU_I                                 'mode de travail du redresseur U(tension)=0, I(intensité)=1
    
    TDetailsPhases(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases
    
    '********** AVEC LES DETAILS POUR AVOIR LA GAMME COMPLETE **********
    TDetailsGammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As EnrDetailsGammesAnodisation 'gamme d'anodisation à executer

    '********** UTILISER UNIQUEMENT EN PRODUCTION **********
    ChoixPosteAnodisation As CHOIX_POSTE_ANODISATION  'choix du poste d'anodisation (imposer au chargement)

End Type

'***************************************************************************************************************************
'                                                                                  CHARGES
'***************************************************************************************************************************

'--- détails des charges (ATTENTION à la correspondance avec le chargement) ---
Public Type DetailsCharges
    NumCommandeInterne As Long                                                        'n° de commande interne = GACLEUNIK
    NumGamme As Long                                                                  'n° de gamme provenant de Clipper
    Naf As Long                                                                       'N° affaire
    TypeReparation As String                                                                              'nombre de réparations (champ texte volontaire)
    CodeClient As String                                                                                      'code du client
    NbrPieces As Double                                                                                        'nombre de pièces
    Designation As String                                                                                    'désignation
    Matiere As String                                                                                           'matières des pièces
    Observations As String                                                                                 'observations
    NumLignesReferencesClient As String                                                         'n° de lignes des références du client correspondant
                                                                                                                           'aux n° de lignes des travaux séparés par un tiret
End Type

'--- détails des fiches de production ---
Public Type DetailsFichesProduction
    
    NumPoste As Integer                                'numéro du poste
    
    DateEntreePoste As Date                         'date d'entrée dans le poste
    DateSortiePoste As Date                          'date de sortie du poste
    DateDebutEgouttage As Date                   'date de début de l'égouttage
    DateFinEgouttage As Date                       'date de début de l'égouttage
    
    TemperatureEnEntree As Single              'température en entrée de bain (si température)
    TemperatureEnSortie As Single               'température en sortie de bain (si température)
    GrapheTemperature As String                  'graphe de la températrure
    
    URedresseur As Single                            'tension redresseur (si redresseur)
    IRedresseur As Single                              'intensité redresseur (si redresseur)
    SensRedresseur As Integer                      'sens du redresseur en fonction du type de redresseur (voir l'énumération correspondante)
    GrapheRedresseur As String                    'graphe du redresseur
    
    AnalyseurEnEntree As Single                  'valeur de l'analyseur en entrée de bain d'anodisation
    AnalyseurEnSortie As Single                   'valeur de l'analyseur en sortie de bain d'anodisation
    GrapheAnalyseur As String                      'graphe de l'analyseur
    
    AlarmesPoste As String                            'alarmes du poste concerné

End Type

'--- états des charges ---
Public Type etatsCharges

    DateEntreeEnLigne As Date                                      'date d'entrée dans la ligne (généralement le chargement)
    DateArriveeAuDechargement As Date                       'date d'arrivée au déchargement
    
    NumBarre As Integer                                                 'numéro de barre
    NumBarreInc As Integer                                                 'numéro de barre incrémental et journalier
    ChargePrioritaire As Boolean                                    'indique qu'il sagit  d'une charge prioritaire
                                                                                       'cette option est validé au chargement
    
    DelaiSupStabilisationChargeSecondes As Integer   'délai supplémentaire de stabilisation de la charge en
                                                                                       'secondes en arrêt au poste pour éviter le mouvement
                                                                                       'pendulaire
    Options1 As Integer                                                   'cette valeur est transmise à l'automate et permet de gérer
                                                                                       'les options 1 (vitesses de montée-descente, etc ...)
                                                                                       'pour certaines charges sur les ponts
    Options2 As Integer                                                   'cette valeur est transmise à l'automate et permet de gérer
                                                                                       'les options 2 (vitesses de montée-descente, etc ...)
                                                                                       'pour certaines charges sur les ponts
   
    
    VitesseHaut As Integer
    VitesseBas As Integer
        
    
    TDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As DetailsCharges 'voir ci-dessus
    
    TGammesAnodisation As EnrGammesAnodisation  'gammes d'anodisation complète pour la production
    PtrZoneGammeAnodisation As Integer                     'pointeur de la zone de la gamme d'anodisation
    
    ModeUouI As MODES_U_OU_I                                 'mode de travail du redresseur U(tension)=0, I(intensité)=1
    
    FinPhase4 As Boolean
    
    
    TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases  'pour renseigner le redresseur en entrant dans le bain
    
    TempsTotalGammeRedresseur As Long                   'temps total de la gamme redresseur en secondes
    
    NbrPostesTraites As Integer                                     'incrémentation de 1 à chaque entrée dans un poste
                                                                                       'sert d'index pour les détails des fiches de production
    TDetailsFichesProduction(1 To NBR_LIGNES_DETAILS_FICHES_PRODUCTION) As DetailsFichesProduction
    
    AlarmesLigne As String                                             'alarmes de la ligne (séparation par -)
    
End Type

'***************************************************************************************************************************
'                                                                CHARGEMENT ET PREVISIONNEL
'***************************************************************************************************************************

'--- chargement ---
Public Type VarChargement
    TDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As DetailsCharges                                                                                                  'voir ci-dessus
    TGammesAnodisation As EnrGammesAnodisation                                                                                                                                              'gammes d'anodisation complète pour la production
    TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases 'pour renseigner le redresseur en entrant dans le bain
End Type

'--- prévisionnel ---
Public Type VarPrevisionnel
    
    ChoixIA As Integer                                                                'meilleur choix pour l'entrée dans la ligne (moteur d'inférence)
    NumCommandeInterne As Long                                         'n° de commande interne
    NumGamme As Long
    Naf As Long
    TypeReparation As String                                                     'nombre de réparations (champ texte volontaire)
    CodeClient As String                                                             'code du client
    NbrPieces As Double                                                               'nombre de pièces
    Designation As String                                                           'désignation
    Observations As String                                                         'observations
    Matiere As String                                                                   'matière des pièces
    NumBarre As Integer                                                             'numéro de barre
    
    NumGammeAnodisation As String                                        'n° de la gamme d'anodisation
    TGammesAnodisation As EnrGammesAnodisation              'gammes d'anodisation complète pour la production
    ChoixPosteAnodisation As CHOIX_POSTE_ANODISATION  'choix du poste d'anodisation

End Type

'--- ligne calcul du prévisionnel ---
Public Type LignePrevisionnel
    
    NumLigne As Integer                                                             'n° de ligne utile pour le prévisionnel
    NumGammeAnodisation As String                                        'n° de la gamme d'anodisation
    TempsPostePrincipalSecondes As Long                              'temps au poste principal en secondes
    
    PassageAnodisation As Boolean                                          'indique un passage dans un des bains d'anodisation
    PassageSpectro As Boolean                                                 'indique un passage dans le bain de spectrocoloration
    PassageOr As Boolean                                                          'indique un passage dans le bain d'or
    PassageNoir As Boolean                                                       'indique un passage dans le bain de noir

End Type

'***************************************************************************************************************************
'                                                               TRACABILITE DE LA PRODUCTION
'***************************************************************************************************************************

'--- enregistrement du type détails des charges de production ---
Public Type EnrDetailsChargesProduction
    
    NumCommandeInterne As Long                            'n° de commande interne
    TypeReparation As String                                        'nombre de réparations (champ texte volontaire)
    DateEntreeEnLigne As Date                                    'date d'entrée dans la ligne (généralement le chargement)
    DateArriveeAuDechargement As Date                     'date d'arrivée au déchargement
    NumBarre As Integer                                               'n° de barre
    NumLigne As Integer                                               'n° de ligne
    CodeClient As String                                                'code du client
    NbrPieces As Double                                                  'nombre de pièces
    Designation As String                                              'désignation
    Matiere As String                                                     'matière des pièces
    NumLignesReferencesClient As String                   'n° de lignes des références du client correspondant
                                                                                     'aux n° de lignes des travaux séparés par un tiret
    
    NumGammeAnodisation As String                          'n° de la gamme d'anodisation
    RefGammeAnodisation As String                            'référence de la gamme d'anodisation
    
    NumFicheProduction As String                                'n° de la fiche de production
    ChargePrioritaire As Boolean                                  'indique qu'il sagit  d'une charge prioritaire
    AlarmesLigne As String                                           'alarmes de la ligne
    ControleColmatage As Integer                                 'contrôle du colmatage (valeur de 0 à 5)
    ControleEpaisseurAnodisation As Integer               'contrôle de l'épaisseur d'anodisation (valeur de 0 à 100)
    ControleColoration As String                                    'contrôle de la coloration (20 caractères)
    ControleObservations As String                               'observations sur les contrôles (50 caractères)

End Type

'--- enregistrement du type détails des gammes de production ---
Public Type EnrDetailsGammesProduction
    NumFicheProduction As String                                     'n° de la fiche de production
    NumLigne As Integer                                                    'n° de ligne
    NumZone As Integer                                                     'n° de la zone
    TempsAuPosteTexte As String                                     'Temps au poste en texte au format HH:MM:SS
    TempsEgouttageTexte As String                                   'Temps d'égouttage en texte au forma MM:SS
    TempsAuPosteSecondes As Long                                'Temps au poste en secondes
    TempsEgouttageSecondes As Integer                          'Temps d'égouttage en secondes
    DecompteDuTempsAuPosteReelSecondes As String  'Représente la différence entre le temps théorique au poste
                                                                                          'et le temps réel passé dans le poste
                                                                                          'un nombre négatif apparait si la charge est resté plus
                                                                                          'longtemps dans le poste que le temps théorique prévu
                                                                                          'ATTENTION variable du type String volontairement
                                                                                          'Si "" alors il n'y a pas eu de temps de décompter
    NumPosteReel As Integer                                            'N° de poste réel utilisé dans la zone (cas des postes multiples)
End Type

'--- enregistrement du type détails des fiches de production ---
Public Type EnrDetailsFichesProduction
    NumFicheProduction As String                  'n° de la fiche de production
    NumLigne As Integer                                 'n° de ligne
    NumPoste As Integer                                 'n° du poste
    DateEntreePoste As Date                          'date d'entrée dans le poste
    DateSortiePoste As Date                           'date de sortie du poste
    DateDebutEgouttage As Date                    'date de début de l'égouttage
    DateFinEgouttage As Date                        'date de début de l'égouttage
    TemperatureEnEntree As Single               'température en entrée de bain (si température)
    TemperatureEnSortie As Single                'température en sortie de bain (si température)
    GrapheTemperature As String                   'graphe de la températrure
    URedresseur As Single                             'tension redresseur (si redresseur)
    IRedresseur As Single                              'intensité redresseur (si redresseur)
    SensRedresseur As Integer                      'sens du redresseur en fonction du type de redresseur (voir l'énumération correspondante)
    GrapheRedresseur As String                    'graphe du redresseur
    AnalyseurEnEntree As Single                   'valeur de l'analyseur en entrée de bain d'anodisation
    AnalyseurEnSortie As Single                    'valeur de l'analyseur en sortie de bain d'anodisation
    GrapheAnalyseur As String                       'graphe de l'analyseur
    AlarmesPoste As String                             'alarmes du poste concerné
End Type

'--- enregistrement du type détails des phases de production ---
Public Type EnrDetailsPhasesProduction
    NumFicheProduction As String                  'n° de la fiche de production
    NumRedresseur As Integer                       'n° du redresseur
    ModeUouI As MODES_U_OU_I                 'mode tension ou intensité
    NumPhase As Integer                                'numéro de la phase
    TempsPhase As Integer                            'temps de la phase
    UPhase As Single                                      'U de la phase
    IPhase As Single                                        'I de la phase
End Type

'***************************************************************************************************************************
'                                                                                  DEFAUTS
'***************************************************************************************************************************

'--- enregistrement du type des défauts ---
Public Type EnrDefauts
    
    NumDefaut As Integer                                          'n° du défaut
    
    SignalerOuiNon As Boolean                                 'indique si le défaut doit être ou pas signaler
    
    GyrophareOuiNon As Boolean                              'indique que le défaut déclenche ou non le gyrophare de la ligne
    
    KlaxonOuiNon As Boolean                                    'indique que le défaut déclenche ou non le klaxon de la ligne
    
    MessageVocalOuiNon As Boolean                       'indique si le défaut déclenche ou non l'envoi d'un message vocal
    
    AfficheurOuiNon As Boolean                                'indique si le défaut concerne l'afficheur à leds rouge
    
    InformationsAPI As String                                     'informations API concernant le défaut
    LibelleDefaut As String                                         'libellé du défaut
    LibelleDefautAfficheur As String                           'libellé du défaut destiné à l'afficheur

    TNumIntervenants(1 To 5) As Integer                   'tableau contenant les numéros (index de personne) des intervenants
    
    '********** NON MEMORISER DANS LA BASE DE DONNEES, UTILISER UNIQUEMENT EN INTERNE **********
    AntiRebondGyrophare As Boolean                      'anti-rebond de signalisation du défaut sur gyrophare
    AntiRebondKlaxon As Boolean                            'anti-rebond de signalisation du défaut sur klaxon
    AntiRebondTraçabiliteAlarmes As Boolean         'anti-rebond de signalisation du défaut dans la table de traçabilité
                                                                                 'des alarmes

End Type

'--- enregistrement du type des intervenants ---
Public Type EnrIntervenants
    NomIntervenant As String
    OrdreAppel As Integer
    ActifOuiNon As Boolean
End Type

'***************************************************************************************************************************
'                                                                  COMMANDES OPERATEUR
'***************************************************************************************************************************

'--- variable des commandes opérateur ---
Public Type VarCommandesOperateur
    TypeCycle As TYPES_CYCLES                           'type de cycle fonction de l'énumération TYPES_CYCLES
    NumPont As Integer                                             'numéro du pont concerné par la commande
    NumPosteDepart As Integer                                'numéro du poste de départ concerné par la commande
    NumPosteArrivee As Integer                               'numéro du poste d'arrivée concerné par la commande
    TempsEgouttageSecondes As Integer                'temps d'égouttage en secondes concerné par la commande
End Type

'***************************************************************************************************************************
'                                                                     MOTEUR D'INFERENCE
'***************************************************************************************************************************

'--- variable de l'ordre de sortie des charges ---
Public Type VarOrdreSortieCharges
    
    NumPoste As Integer                                                                      'numéro du poste
    NumCharge As Integer                                                                    'numéro de charge au poste
    
    NumPosteArrivee As Integer                                                          'numéro du poste d'arrivée prévu
    NumChargePosteArrivee As Integer                                              'numéro de charge au poste d'arrivée
    
    DecompteDuTempsAuPosteReelSecondes As String                    'décompte du temps au poste réel en secondes
    Condamnation As Boolean                                                             'TRUE=Poste condamné
                                                                                                            'FALSE=Poste en fonctionnement normal
    NumPont As Integer                                                                        'numéro du pont théorique de la prémisse choisi
                                                                                                            'pour le futur mouvement lorsque le temps sera à 0
                                                                                                    
End Type

'--- variable des informations sur les postes d'anodisation ---
Public Type VarInformationsPostesAnodisation
    NumCharge As Integer                                                                    'numéro de charge au poste
    DecompteDuTempsAuPosteReelSecondes As String                    'décompte du temps au poste réel en secondes
    Condamnation As Boolean                                                             'TRUE=Poste condamné
                                                                                                            'FALSE=Poste en fonctionnement normal
End Type

'--- variable du moteur d'inférence ---
Public Type VarMoteurInference
    
    '--- chargement ---
    ProchainNumPosteChargementSiAnodisationC13Impose As Integer     'prochain n° de poste au chargement si le poste
                                                                                                                     'd'anodisation C13 est imposé dans la gamme
    ProchainNumPosteChargementSiAnodisationC14Impose As Integer     'prochain n° de poste au chargement si le poste
                                                                                                                     'd'anodisation C14 est imposé dans la gamme
    ProchainNumPosteChargementSiAnodisationC15Impose As Integer     'prochain n° de poste au chargement si le poste
                                                                                                                     'd'anodisation C15 est imposé dans la gamme
    ProchainNumPosteChargementSiAnodisationC16Impose As Integer     'prochain n° de poste au chargement si le poste
                                                                                                                     'd'anodisation C16 est imposé dans la gamme
    
    ProchainNumPosteChargementSiAnodisationAutomatique As Integer   'prochain n° de poste au chargement si le poste
                                                                                                                     'd'anodisation est automatique dans la gamme
    ProchainNumPosteChargement As Integer                                             'indique le prochain numéro de poste ou se fera
                                                                                                                     'le chargement
    
    '--- ordre de sortie des charges dans la ligne ---
    TOrdreSortieCharges(1 To CHARGES.C_NUM_MAXI) As VarOrdreSortieCharges
                                                                                                                    'tableau contenant l'ordre de sortie des charges
                                                                                                                    'ce tableau est trié directement du temps le plus
                                                                                                                    'court au temps le plus long
                                                                                                                    'CHARGES.C_NUM_MAXI correspondant au
                                                                                                                    'nombre de charges maxi dans la ligne
    '--- ordre de sortie pour les ponts ---
    TOrdreSortiePonts(PONTS.P_1 To PONTS.P_2, 1 To CHARGES.C_NUM_MAXI) As VarOrdreSortieCharges
                                                                                                                    'tableau contenant l'ordre de sortie des charges
                                                                                                                    'ce tableau est trié directement du temps le plus
                                                                                                                    'court au temps le plus long
                                                                                                                    'CHARGES.C_NUM_MAXI correspondant au
                                                                                                                    'nombre de charges maxi dans la ligne
    
    '--- informations complètes sur les postes d'anodisation ---
    'sert à connaitre le moment opportum pour l'entrée d'une charge dans la ligne
    TInformationsPostesAnodisation(POSTES.P_C13 To POSTES.P_C16) As VarInformationsPostesAnodisation
                                                                                                                   'tableau contenant les informations sur les
                                                                                                                   'postes d'anodisation
End Type

'***************************************************************************************************************************
'                                                                  TRACABILITE DES ALARMES
'***************************************************************************************************************************

'--- enregistrement du type de la traçabilité des alarmes ---
Public Type EnrTraçabiliteAlarmes
    ClePrimaire As Long                         'clé primaire
    NumDefaut As Integer                       'n° du défaut
    DateDetectionDefaut As Date           'date de détection du défaut
End Type

'***************************************************************************************************************************
'                                                                        JOURNEES TYPES
'***************************************************************************************************************************

'--- variable d'un cycle de 24 heures pour les journées types ---
Public Type VarCycle24HJourneesTypes
    TypeDeJournee As JOURNEES_TYPES                                               'type de journée par cuve
    TTopsDebutPompe(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de début d'un cycle pompe
    TTopsFinPompe(1 To NBR_TOPS_POSSIBLES) As String * 14           'tops de fin d'un cycle pompe
    TCyclesPompe(1 To NBR_TOPS_POSSIBLES) As Integer                   'cycles de la pompe
    TTopsDebutChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14  'tops de début d'un cycle chauffage
    TTopsFinChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de fin d'un cycle chauffage
    TModesChauffage(1 To NBR_TOPS_POSSIBLES) As Integer               'modes du chauffage
End Type

'***************************************************************************************************************************
'                                                                PROGRAMMATEUR CYCLIQUE
'***************************************************************************************************************************

'--- variable d'un cycle de 24 heures pour le programmateur cyclique ---
Public Type VarCycle24HProgCyclique
    TypeDeJournee As JOURNEES_TYPES                                               'type de journée par cuve
    TTopsDebutPompe(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de début d'un cycle pompe
    TTopsFinPompe(1 To NBR_TOPS_POSSIBLES) As String * 14            'tops de fin d'un cycle pompe
    TCyclesPompe(1 To NBR_TOPS_POSSIBLES) As Integer                    'cycles de la pompe
    TTopsDebutChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14   'tops de début d'un cycle chauffage
    TTopsFinChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de fin d'un cycle chauffage
    TModesChauffage(1 To NBR_TOPS_POSSIBLES) As Integer               'modes du chauffage
End Type

'***************************************************************************************************************************
'                                                                               ANNEXES
'***************************************************************************************************************************

'--- numéros des défauts ---
'Public Type NumDefautsAnnexes
    
'End Type

'--- états des annexes ---
Public Type EtatsAnnexes
    
    'TNumDefauts As numdefautsAnnexes
    
    API_ChangementsEVBrillantage As Boolean                                     'TRUE=Indique un changement (hors états)
    API_EtatsEVBrillantage As String * 16                                                'retour en binaire du mot de l'API contenant les états
    EtatsEVBrillantage As ETATS_EV_BRILLANTAGE                              'fonction de l'énumération
    PeriodiciteEVBrillantage As Integer                                                    'périodicité de mise en marche de l'électro-vanne d'air dans le bain de brillantage
    TempsMarcheEVBrillantage As Integer                                              'temps de marche de l'électro-vanne d'air dans le bain de brillantage
    ModeEVBrillantage As MODES_EV_BRILLANTAGE                           'fonction de l'énumération

    API_ChangementsEVEauLigne As Boolean                                        'TRUE=Indique un changement (hors états)
    EtatsEVEauLigne As ETATS_EV_EAU_LIGNE                                     'fonction de l'énumération
    ModeEVEauLigne As MODES_EV_EAU_LIGNE                                  'fonction de l'énumération
    
    API_ChangementsCompresseurP1 As Boolean                                  'TRUE=Indique un changement (hors états)
    EtatsCompresseurP1 As ETATS_COMPRESSEURS_PONTS              'fonction de l'énumération
    ModeCompresseurP1 As MODES_COMPRESSEURS_PONTS           'fonction de l'énumération
    
    API_ChangementsEclairageP1 As Boolean                                         'TRUE=Indique un changement (hors états)
    EtatsEclairageP1 As ETATS_ECLAIRAGE_PONTS                               'fonction de l'énumération
    ModeEclairageP1 As MODES_ECLAIRAGE_PONTS                            'fonction de l'énumération
    
    API_ChangementsCompresseurP2 As Boolean                                  'TRUE=Indique un changement (hors états)
    EtatsCompresseurP2 As ETATS_COMPRESSEURS_PONTS              'fonction de l'énumération
    ModeCompresseurP2 As MODES_COMPRESSEURS_PONTS           'fonction de l'énumération
    
    API_ChangementsEclairageP2 As Boolean                                         'TRUE=Indique un changement (hors états)
    EtatsEclairageP2 As ETATS_ECLAIRAGE_PONTS                               'fonction de l'énumération
    ModeEclairageP2 As MODES_ECLAIRAGE_PONTS                             'fonction de l'énumération
    
    TAPI_ChangementsNiveauxRetentions(NIVEAUX_RETENTIONS.NR_STOCKAGE_STATION To NIVEAUX_RETENTIONS.NR_LAVEUR) As Boolean                                          'TRUE=Indique un changement (hors états)
    TEtatsNiveauxRetentions(NIVEAUX_RETENTIONS.NR_STOCKAGE_STATION To NIVEAUX_RETENTIONS.NR_LAVEUR) As ETATS_NIVEAUX_RETENTIONS                          'fonction de l'énumération

End Type

'***************************************************************************************************************************
'                                                                   COMMANDES INTERNE
'***************************************************************************************************************************

'--- enregistrement du type "CommandesInterne" ---
Public Type EnrCommandesInterne
    
    NumCommandeInterne As Long             'n° de commande interne
    CodeClient As String                                'Code du client
    Designation As String                               'désignation
    NbrPieces As Double                                   'nombre de pièces
    Matiere As String                                      'matière des pièces

End Type

'***************************************************************************************************************************
'                                                                         PHASES DE CLIPPER
'***************************************************************************************************************************

'--- enregistrement du type "PhasesClipper" ---
Public Type EnrPhasesClipper
    GaCLeUnik  As Long                          'clé unique pour rechercher une affaire
    CoFrais As String                               'centre de frais
    CoCli As String                                   'code client
    NomClient As String                           'nom du client
    Piece As String                                  'référence de la pièce
    QteAf As Double                                    'quantité de pièces
    Desa1 As String                                 'désignation de la pièce
    DateLance As String                          'date de lancement format texte JJ/MM/AAAA
    Matiere As String                               'matière de la pièce
    GamObs As String                             'observations
    NumGamme As Long
    Naf As Long
End Type

'***************************************************************************************************************************
'                                                               GRAPHES DE PRODUCTION
'***************************************************************************************************************************

'--- traçabilité ---
Public Type Traçabilite
    DateDuPoint As Date                                  'date du point (construction de l'axe des x)
    NumPhase As Byte                                     'numéro de la phase
    EtatRedresseur As Byte                             'état du redresseur (0=Arrêt, 1=Pause, 2= Marche, 3=Défaut)
    Tension As Byte                                          'tension mesurée (valeur en décimale fois 10)
    Intensite As Integer                                     'intensité mesurée
    Temperature As Integer                               'température mesurée (valeur en décimale fois 10)
End Type

'--- renseignements d'un graphe ---
Public Type RenseignementsGraphe
    NumFicheProduction As String                    'n° de la fiche de production
    DateEntreeEnLigne As Date                        'date d'entrée dans la ligne (généralement le chargement)
    NumRedresseur As Integer                         'numéro du redresseur
End Type

'***************************************************************************************************************************
'                                                                           ETUVE en C33/C34
'***************************************************************************************************************************

'--- entrées API de l'étuve C33/C34 ---
Public Type EntreesEtuveC33C34_OLD
    
    DesequilibrePhases As Boolean                             'déséquilibre des phases

    SurchauffeElementChauffant1 As Boolean              'surchauffe de l'élément chauffant 1
    SurchauffeElementChauffant2 As Boolean              'surchauffe de l'élément chauffant 2
    SurchauffeElementChauffant3 As Boolean              'surchauffe de l'élément chauffant 3
    
    DefautMarcheVentilateur As Boolean                      'défaut de marche du ventilateur

End Type

'--- sorties API de l'étuve C33/C34 ---
'Public Type SortiesEtuveC33C34

'End Type

'--- numéros des défauts ---
Public Type NumDefautsEtuveC33C344_OLD
    
    NumDefautDesequilibrePhases As Integer

    NumDefautSurchauffeElementChauffant1 As Integer
    NumDefautSurchauffeElementChauffant2 As Integer
    NumDefautSurchauffeElementChauffant3 As Integer
    
    NumDefautDefautMarcheVentilateur As Integer
    
End Type

Public Type EtatsEtuveC33C344_OLD
    
    TNumDefauts As NumDefautsEtuveC33C344_OLD                             'numéros des défauts
    
    ManuelAutomatique As Boolean                                                'manuel / automatique de l'étuve
    
    DepartCycle As Boolean                                                            'départ cycle
    CycleEnCours As Boolean                                                         'cycle en cours
    CycleTermine As Boolean                                                         'cycle terminé
    
End Type

'***************************************************************************************************************************
'                                                                                DIVERS
'***************************************************************************************************************************

'--- manipulations dans la fenêtre gestion de la régulation ---
Public Type ManipsGestionRegulation
    AppareillageConcerne As Boolean           'FALSE=pompe, TRUE=chauffage
    CyclesPompe As Integer
    ModesChauffage As Integer
End Type

'--- manipulations dans la fenêtre du programmateur cyclique ---
Public Type ManipsProgCyclique
    AppareillageConcerne As Boolean           'FALSE=pompe, TRUE=chauffage
    CyclesPompe As Integer
    ModesChauffage As Integer
End Type

Attribute VB_Name = "MPTypes"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DES TYPES PUBLIQUES
' Nom                    : MPTypes.bas
' Date de cr�ation : 26/03/1999
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
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
    NumAction As Long                          'N� de l'action
    CodeAction As String                        'Code de l'action
    LibelleAction As String                     'libell� de l'action
    ParametreOuiNon As Boolean         'Indique si l'action a un param�tre
    LibelleParametre As String              'libell� du param�tre
End Type

'***************************************************************************************************************************
'                                                                          CYCLE D'UN PONT
'***************************************************************************************************************************

'--- cycle d'un pont ---
Public Type VarCyclePont
    TActions(0 To NBR_LIGNES_CYCLES_PONTS) As EnrActions  'tableau contenant le cycle du pont
    PtrActions As Integer                                                                   'pointeur des actions du pont
    NumAction As Long                                                                      'num�ro de l'action en cours du pont
End Type

'***************************************************************************************************************************
'                                                                   TEMPS DE MOUVEMENTS
'***************************************************************************************************************************

'--- variable d'analyse des mouvements ---
Public Type VarAnalyseMouvements
    EtatMouvement As ETATS_MOUVEMENTS          '0 = Pas de mouvement (�tat de repos avant le d�but du mouvement)
                                                                                  '1 = Le mouvement est en cours
                                                                                  '2 = Fin du mouvement
                                                                                  '3 = Calcul du temps r�el du mouvement puis remise � 0
    DateDebutMouvement As Date                            'date de d�but d'un mouvement
    DateFinMouvement As Date                                 'date de fin d'un mouvement
End Type

'***************************************************************************************************************************
'                                                                               PREMISSES
'***************************************************************************************************************************

'--- variable du type des pr�misses pour la cr�ation du tableau des pr�misses ---
Public Type VarPremisses
    NumPont As Integer                                              'n� du pont concern� d�fini comme r�gle au d�part
    NumPontIA As Integer                                           'n� du pont choisi par le moteur d'inf�rence
    PremisseCodee As String                                     'pr�misse cod�e
    PremisseDecodee As String                                 'pr�misse d�cod�e
    TempsCycleSecondes As Long                            'temps du cycle en secondes
End Type

'--- variable du type des pr�misses pour le calcul du temps de cycle ---
Public Type VarPremissesTempsCycle
    NumAction As Integer                      'n� de l'action
    CodeAction As String                       'code de l'action
    ParametreOuiNon As Boolean        'param�tre oui ou non
    Parametre As String                        'param�tre en fonction de l'action
    LibelleAction As String                    'libell� de l'action
End Type

'***************************************************************************************************************************
'                                                                             REDRESSEURS
'***************************************************************************************************************************

'--- d�tails des phases U et I pour renseigner le redresseur en entrant dans le bain ---
Public Type DetailsPhases
    TempsPhase As Integer                                                            'temps de la phase
    UPhase As Single                                                                     'tension de la phase
    IPhase As Single                                                                       'intensit� de la phase
End Type

'--- enregistrement du type redresseurs ---
Public Type EnrRedresseurs
    NumRedresseur As Integer                                                       'N� du redresseur
    NomRedresseur As String                                                         'nom du redresseur
    LibelleRedresseur As String                                                      'libell� du redresseur
    UMaxiRedresseur As Single                                                      'tension maximale donn�e par le constructeur du redresseur
    IMaxiRedresseur As Single                                                       'courant maximum donn� par le constructeur du redresseur
End Type

'--- entr�es des redresseurs ---
Public Type EntreesRedresseurs

    M_DefautGeneral As Boolean                                                    'd�faut g�n�ral
    
    M_DelaiTropLongMarcheRedresseur As Boolean                     'd�lai trop long de mise en marche du redresseur (contr�le sur l'intensit�)

    M_IntensiteNonAtteinte As Boolean                                           'intensit� demand�e non atteinte
    
    M_IntensiteInstable As Boolean                                                 'intensit� instable en cours de fonctionnement sur la phase 4
    
End Type

'--- num�ros des d�fauts ---
Public Type NumDefautsRedresseurs
    
    NumDefautDefautGeneral As Integer                                         'num�ro de d�faut du d�faut g�n�ral
    
    NumDefautDelaiTropLongMarcheRedresseur As Integer          'num�ro de d�faut du d�lai trop long de mise en marche du redresseur (contr�le sur l'intensit�)
   
    NumDefautIntensiteNonAtteinte As Integer                                'num�ro de d�faut du intensit� demand�e non atteinte
   
    NumDefautIntensiteInstable As Integer                                      'num�ro de d�faut du intensit� instable en cours de fonctionnement sur la phase 4
   
End Type

'--- �tats des redresseurs ---
Public Type EtatsRedresseurs
    
    TNumDefauts As NumDefautsRedresseurs                                'num�ros des d�fauts
    
    TEntreesAPI As EntreesRedresseurs                                          'entr�es des redresseurs (automate)
    
    DefinitionRedresseur As EnrRedresseurs                                  'd�finition d'un redresseur
    
    ModeRedresseur As MODES_REDRESSEUR                             'mode du redresseur
    EtatRedresseur As ETATS_REDRESSEUR                                  '�tat �lectrique du redresseur
    SensRedresseur As SENS_REDRESSEUR                                 'sens du redresseur
    
    EtatsMarcheArret As String                                                          '�tats de marche et d'arr�t
    Etats1 As String                                                                           '�tats 1 du redresseur
    Etats2 As String                                                                           '�tats 2 du redresseur
    
    DemandesDuPC As Integer                                                         'demande du PC
    RetoursVersPC As Integer                                                           'retours vers PC (les valeurs retourn�es doivent correspondre aux valeurs demand�es)
    IDemandePC As Integer                                                               'Intensit� demand� par le PC
    
    NumCharge As Integer                                                                 'num�ro de la charge trait� par le redresseur
    
    NumPhaseEnCours As PHASES_GAMMES_REDRESSEURS      'num�ro de la phase en cours (gamme d'anodisation en automatique)
    TempsPhaseEnCours As Integer                                                  'temps de la phase en cours
    TempsEcoulePhaseEnCours As Integer                                       'temps �coul� de la phase en cours
    
    DebutCycle As Boolean                                                                'TRUE = Indique le d�but d'un cycle, FALSE = Pas de cycle en cours
    ControleFinCycle As Boolean                                                       'TRUE = Indique la fin d'un cycle, FALSE = Cycle en cours
    
    TempsAjouteSurIFaible As Integer                                                'temps ajout� sur une intensit� plus faible que pr�vue (panne de redresseur)
    TempsTotalCycle As Integer                                                         'temps total du cycle en secondes lu dans l'automate
    TempsRestantCycle As Integer                                                     'temps restant du cycle en secondes lu dans l'automate
    TempsTotalise As Integer                                                             'temps totalis�
    
    ModeUouI As MODES_U_OU_I                                                     'mode de travail du redresseur U(tension)=0, I(intensit�)=1
    
    TDetailsPhases(PHASES_GAMMES_REDRESSEURS.PH_T1 To _
                              PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases    'd�tails des phases
    
    U As Single                                                                                   'tension mesur�e en volts
    I As Integer                                                                                    'courant mesur�e en amp�res
    
    ConsigneU As Single                                                                    'consigne en tension
    ConsigneI As Integer                                                                    'consigne en intensit�

    Ah As Single                                                                                  'Ah calcul�e
    
End Type

'***************************************************************************************************************************
'                                                                          ETATS DE LA LIGNE
'***************************************************************************************************************************

'--- num�ros des d�fauts ---
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
    
    TNumDefauts As NumDefautsLigne  'num�ros des d�fauts
    
    MarcheGenerale As Boolean          'marche g�n�rale de la ligne
    ArretGeneral As Boolean                'arr�t g�n�ral de la ligne
    ArretUrgence As Boolean               'arr�t d'urgence
    PortillonsLigneVie As Boolean       'portillons et ligne de vie
    StopLigne As Boolean                    'stop ligne
    
    ArretUrgenceP1 As Boolean           'arr�t d'urgence du pont 1
    ArretUrgenceP2 As Boolean           'arr�t d'urgence du pont 2
    
    SecuriteP1 As Boolean                   's�curit� du pont 1
    SecuriteP2 As Boolean                   's�curit� du pont 2
    
    ManqueTension As Boolean           'manque de tension
    ManqueAir As Boolean                    'manque d'air
    AcquittementsDefauts As Boolean  'pour l'acquittement des d�fauts
    FrontMontantDefauts As Boolean    'front montant des d�fauts

End Type

'***************************************************************************************************************************
'                                                                                   PONTS
'***************************************************************************************************************************

'--- param�tres des cycles des ponts ---
Public Type ParametresCyclesPonts
    NumPosteDepart As Integer                                    'poste de d�part d'un cycle d'un pont
    NumPosteArrivee As Integer                                   'poste d'arriv�e d'un cycle d'un pont
    TypeCycle As TYPES_CYCLES                               'd�placement ou transfert de charge
    DelaiSupStabilisationChargeSecondes As Integer 'd�lai suppl�mentaire de stabilisation de la charge
    TempsEgouttageSecondes As Integer                    'temps d'�gouttage en secondes
End Type

'--- cycles des ponts ---
Public Type CyclesPonts
    NumAction As Integer                   'N� de l'action
    Parametre As String                     'Param�tre en fonction de l'action
    EtatParametre As String               'Etat des param�tres (indique le temps qu'il reste ou autre)
End Type

'--- variable du type pointeur de l'action et action en cours donn�es par l'API ---
Public Type VarPtrEtActionEnCoursAPI
    PtrAction As Integer                      'pointeur de l'action
    NumAction As Integer                   'N� de l'action
    Parametre As Integer                   'param�tre
End Type

'--- entr�es des ponts API ---
Public Type EntreesPontsAPI
    
    M_MoteurTourneTrlPont As Boolean                                      'le moteur tourne de la translation gauche du pont

    M_MoteurTourneLevPont As Boolean                                    'le moteur tourne du levage du pont
    
    M_MarquageAxeTrL As Boolean                                            'marquage axe de la translation
    M_MarquagePVTrL As Boolean                                              'marquage de la petite vitesse de la translation
    M_MarquageMVTrL As Boolean                                             'marquage de la moyenne vitesse de la translation
    M_MarquageArretTrL As Boolean                                          'marquage d'arr�t de la translation
    
    M_MarquageAxeLev As Boolean                                           'marquage axe du levage
    
    M_MemDemandeIsocentrage As Boolean                             'm�moire de demande d'isocentrage apr�s un glissement important
    
    M_DefautPresencePicece  As Boolean                                  'indique poste occup� sur Cellule
    
    M_MarquageAxeLevPont As Boolean                                    'marquage axe du levage du pont
    M_ErreurPointeur As Boolean                                                'erreur sur le pointeur des actions ou sur le code des actions
    
    M_AccrochesEnHaut As Boolean
    M_AccrochesEnBas As Boolean
    
    E_NiveauHaut As Boolean
    E_NiveauIntermediaire As Boolean
    E_NiveauBas As Boolean
    
    M_DefautVariateurTrlPont As Boolean                                   'd�faut variateur de la translation gauche du pont
    M_AxeNonReferenceTrlPont As Boolean                               'axe non r�f�renc� de la translation gauche du pont
    M_SurcourseTrlAvant As Boolean                                          'surcourse de la translation avant
    M_SurcourseTrlArriere As Boolean                                        'surcourse de la translation arri�re
    M_DefautVariateurLevPont As Boolean                                 'd�faut variateur du levage du pont
    M_AxeNonReferenceLevPont As Boolean                              'axe non r�f�renc� du levage du pont
    M_SurcourseLevBas As Boolean                                           'surcourse levage bas
    M_SurcourseLevHaut As Boolean                                          'surcourse levage haut
    
    M_DelaiTropLongDescenteAccroches As Boolean
    M_DelaiTropLongMonteeAccroches As Boolean
    
End Type

'--- sorties des ponts API ---
Public Type SortiesPontsAPI
    
    A_AntiCollision As Boolean                                                    'sortie API du surcourse avant pour le pont 1 et arri�re du pont 2
    
    S_EVMonteeAccroches As Boolean                                       '�lectro-vanne de mont�e des accroches (lib�ration de la charge)
    S_EVDescenteAccroches As Boolean                                    '�lectro-vanne de descente des accroches (prise de la charge)

End Type

'--- temps des mouvements ---
'tous les temps sont en secondes
Public Type TempsMouvementsPonts
    
    TTempsTranslation(POSTES.P_CHGT_1 To DERNIER_POSTE, POSTES.P_CHGT_1 To DERNIER_POSTE) As Single
                                                                                          'temps de d�placement du pont d'un poste de d�part vers
                                                                                          'un poste d'arriv�e
    
    TempsAccrochesChargeVersHaut As Single               'temps accroches charges vers le haut
    TempsAccrochesChargeVersBas As Single                'temps accroches charges vers le bas
    
    TempsDescenteHautVersBas As Single                      'temps en descente du niveau haut vers le niveau bas
    TempsDescenteIntermediaireVersBas As Single        'temps en descente du niveau interm�diaire vers le niveau bas

    TempsMonteeBasVersIntermediaire As Single           'temps en mont�e du niveau bas vers le niveau interm�diaire
    TempsMonteeBasVersHaut As Single                         'temps en mont�e du niveau bas vers le niveau haut

End Type

'--- num�ros des d�fauts ---
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

'--- �tats des ponts ---
Public Type EtatsPonts
    
    ModePont As MODES_PONTS                 'mode du pont de maintenance � automatique
    TypeSequence As TYPES_SEQUENCES 'type de s�quence du pont (inconnu, cyclique, al�atoire)
    ControleParOperateur As Boolean           'FALSE=contr�le par IA, TRUE=contr�le par op�rateur (uniquement en mode automatique)
    
    TypesAffichagesCyclesPonts As Boolean    'pour l'affichage des cycles de ponts
                                                                           'FALSE = Cycle actuel, TRUE = Prochain cycle

    TParametresCyclesPonts(CYCLES.C_ACTUEL To CYCLES.C_PROCHAIN) As ParametresCyclesPonts
                                                                           'param�tres des cycles des ponts (poste de d�part et arriv�e, etc ...)
    TCyclesPonts(CYCLES.C_ACTUEL To CYCLES.C_PROCHAIN, 1 To NBR_LIGNES_CYCLES_PONTS) As CyclesPonts
                                                                     'cycles des ponts

    PtrEtActionEnCoursAPI As VarPtrEtActionEnCoursAPI
                                                                     'pointeur de l'action et action en cours retourn�e par l'automate
                                                                     '(ne concerne que le cycle actuel)
    
    TEntreesAPI As EntreesPontsAPI             'entr�es des ponts (automate)
    TSortiesAPI As SortiesPontsAPI               'sorties des ponts (automate)
    
    TTempsMouvements As TempsMouvementsPonts    'pour la m�morisation des temps de mouvements
    
    TNumDefauts As NumDefautsPonts          'num�ros de d�fauts
    UnDefautAuMoinsSignale As Boolean      'un d�faut au moins est signal�
    
    PositionActuelleLaserTrlPont As Long      'position actuel du laser de la translation gauche
    PositionCibleLaserTrlPont As Long           'position cible du laser de la translation gauche
    
    PositionActuelleCodeurTrlPont As Long    'position actuelle du codeur de la translation gauche
    PositionCibleCodeurTrlPont As Long        'position cible du codeur de la translation gauche
    
    PositionActuelleCodeurLevPont As Long  'position actuelle du codeur de levage
    PositionCibleCodeurLevPont As Long      'position cible du codeur de levage
     
    PosteActuel As POSTES                           'num�ro du poste actuel
    PosteDestination As POSTES                  'num�ro du poste de destination
    SensX As SENS_X                                    '1=sens avant, -1=sens arri�re, 0=arr�t au poste

    NiveauActuel As NIVEAUX_PONTS           'num�ro du niveau actuel (de 201 � 215)
    NiveauDestination As NIVEAUX_PONTS  'num�ro du niveau de destination
    SensY As SENS_Y                                    '1=sens mont�e, -1=sens descente, 0=arr�t au niveau

    EtatsAccrochesCharge As ETATS_ACCROCHES    '0 = Accroches de la charge en haut, 1 = Accroches de la charge en bas

    PoidsSouleve As Single                            'poids soulev�

    Condamnation As Boolean                        'TRUE=Pont condamn�
    NumCharge As Integer                              'N� de la charge sur le pont
    OptionsGamme1 As Integer                      'options de la gamme partie 1
    OptionsGamme2 As Integer                      'options de la gamme partie 2
    
    Alarmes As String                                     'N� des alarmes du pont (disjonction ou s�quence non effectu�e)

End Type

'***************************************************************************************************************************
'                                                                                   POSTES
'***************************************************************************************************************************

'--- enregistrement du type postes ---
Public Type EnrPostes
    
    NumPoste As Integer                                    'num�ro du poste dans la ligne
    NomPoste As String                                      'nom du poste dans la ligne
    LibellePoste As String                                  'libell� complet du poste dans la ligne
    
    AvecTemps As Boolean                               'FALSE=Pas de temps au poste (chargement, d�chargement ...)
                                                                         'TRUE=Avec un temps obligatoire (cas des bains)
    
    RespectTempsObligatoire As Boolean        'FALSE=Le temps de bain peut admettre d'�tre d�pass�
                                                                         'TRUE=Le temps de bain doit �tre respect�
    
    AvecEgouttage As Boolean                          'FALSE=Pas d'�gouttage au poste, TRUE=Avec �gouttage au poste
    
    PresenceCouvercles As Boolean                 'FALSE=Pas de couvercles, TRUE=Pr�sence de couvercles
    PresenceRedresseur As Boolean                'FALSE=Pas de redresseur, TRUE=Pr�sence d'un redresseur
    PresenceAgitationBain As Boolean              'FALSE=Pas d'agitation du bain, TRUE=Pr�sence d'une agitation du bain
    
    XAxePosteLigne As Long                              'X de l'axe de poste dans la ligne (valeur du lecteur laser)
    
    XAxePosteSynoptique As Long                     'X de l'axe de poste dans le synoptique
    
    XInferieurPosteSynoptique As Long             'X inf�rieur du rectangle limitant le poste dans le synoptique
    YInferieurPosteSynoptique As Long             'Y inf�rieur du rectangle limitant le poste dans le synoptique
    XSuperieurPosteSynoptique As Long           'X sup�rieur du rectangle limitant le poste dans le synoptique
    YSuperieurPosteSynoptique As Long           'X sup�rieur du rectangle limitant le poste dans le synoptique
    
    XInferieurLibellePosteSynoptique As Long   'X inf�rieur du rectangle limitant le libell� du poste dans le synoptique
    YInferieurLibellePosteSynoptique As Long   'Y inf�rieur du rectangle limitant le libell� du poste dans le synoptique
    XSuperieurLibellePosteSynoptique As Long 'X sup�rieur du rectangle limitant le libell� du poste dans le synoptique
    YSuperieurLibellePosteSynoptique As Long 'X sup�rieur du rectangle limitant le libell� du poste dans le synoptique

End Type

'--- �tats des postes ---
Public Type EtatsPostes
    
    DefinitionPoste As EnrPostes                                       'd�finition du poste
    
    Condamnation As Boolean                                            'TRUE=Poste condamn�
                                                                                           '           pas de prise ni de d�pose de charge autoris�es
                                                                                           'FALSE=Poste en fonctionnement normal
                                                                                           '           prise et d�pose de charge autoris�es
    
    NumCharge As Integer                                                  'N� de la charge dans le poste
    
    EtatsChariots As ETATS_CHARIOTS                             '�tats des chariots
    
    Alarmes As String                                                          'N� des alarmes du poste (disjonction, etc...)
    
End Type

'***************************************************************************************************************************
'                                                               GESTION DES TEMPERATURES
'***************************************************************************************************************************

'--- variable de gestion des temp�ratures ---
Public Type VarTemperatures
    TempActuelle As Single                     'temp�rature actuelle de la cuve (valeur retourn�e par l'automate)
    TempVeille As Single                         'temp�rature de veille (op�rateur)
    TempProduction As Single                 'temp�rature de production normale (op�rateur)
    EcartInferieurRegul As Single            '�cart inf�rieur de r�gulation (op�rateur)
    EcartSuperieurRegul As Single          '�cart sup�rieur de r�gulation (op�rateur)
    EcartInferieurAlarme As Single          '�cart inf�rieur d'alarme (op�rateur)
    EcartSuperieurAlarme As Single        '�cart sup�rieur d'alarme (op�rateur)
End Type

'***************************************************************************************************************************
'                                                                                    CUVES
'***************************************************************************************************************************

'--- enregistrement du type cuves ---
Public Type EnrCuves
    
    NumCuve As Integer                                                         'num�ro de la cuve dans la ligne
    NomCuve As String                                                           'nom de la cuve dans la ligne
    LibelleCuve As String                                                        'libell� complet de la cuve dans la ligne
   
    GestionAPI As Boolean                                                      'indique si la cuve est g�r�e par l'automate
    
    PresencePompe As Boolean                                             'FALSE=Pas de pompe, TRUE=Pr�sence d'une pompe
    NbrChauffages As Integer                                                  'nombre de chauffages (0=cuve non chauff�e)
    PresenceRefroidissementBain As Boolean                      'FALSE=Pas de refroidissement du bain, TRUE=Refroidissement du bain
    PresenceNiveauBas As Boolean                                      'FALSE=Pas de niveau bas, TRUE=Pr�sence d'un niveau bas
    PresenceNiveauHaut As Boolean                                     'FALSE=Pas de niveau haut, TRUE=Pr�sence d'un niveau haut
    PresenceEVEau As Boolean                                              'FALSE=Pas d'�lectro-vanne d'eau, TRUE=Pr�sence d'une �lectro-vanne d'eau
    PresenceAnalyseurAnodisation As Boolean                     'FALSE=Pas d'analyseur d'anodisation, TRUE=Pr�sence d'un analyseur d'anodisation

End Type

'--- caract�ristiques g�n�rales de toutes les cuves ---
Public Type CaracteristiquesCuves
    
    DefinitionCuve As EnrCuves                 'd�finition d'une cuve

End Type

'--- entr�es des cuves API ---
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

'--- num�ros des d�fauts ---
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

'--- �tats des cuves g�r�es par l'automate ---
Public Type EtatsCuves
    
    DefinitionCuve As EnrCuves                                                                  'd�finition d'une cuve
    IndexAutomate As Integer
    API_Changements  As Boolean                                                             'TRUE=Indique un changement (hors �tats)
    API_Etats_1 As String * 16                                                                     'retour en binaire du mot de l'API contenant les �tats 1
    API_Etats_2 As String * 16                                                                     'retour en binaire du mot de l'API contenant les �tats 2
    
    TEntreesAPI As EntreesCuvesAPI                                                          'entr�es des cuves (automate)
    TSortiesAPI As SortiesCuvesAPI                                                            'sorties des cuves (automate)
    
    TTempsMouvements As TempsMouvementsCuves                             'pour la m�morisation des temps de mouvements
    
    TNumDefauts As NumDefautsCuves                                                     'num�ros de d�fauts
    UnDefautAuMoinsSignale As Boolean                                                   'un d�faut au moins est signal�
    ListeNumDefautsSiCharge As String                                                     'liste des num�ros de d�fauts en pr�sence d'une charge
                                                                                                                   'si pas de charge alors cette variable est vide
     
    API_ModeProduction As MODES_PRODUCTION                                   'retour API du mode de production
    ModeProduction As MODES_PRODUCTION                                          'd�terminer par le programmateur cyclique
    
    ModeRegulation As MODES_REGULATION                                          'mode de la r�gulation
    
    EtatsChauffage As ETATS_CHAUFFAGES                                              '�tats d'un chauffage
    EtatsRefroidissementBain As ETATS_REFROIDISSEMENT_BAIN        '�tats du froidissement d'un bain
    
    API_ModePompe As MODES_POMPES                                                 'retour API du mode de la pompe
    API_CyclePompe As CYCLES_POMPES                                                'retour API du cycle de la pompe
    ModePompe As MODES_POMPES                                                        'mode de la pompe d�terminer par l'op�rateur
    CyclePompe As CYCLES_POMPES                                                       'cycle de la pompe d�terminer par le programmateur cyclique
    EtatsPompe As ETATS_POMPES                                                           '�tats de la pompe

    EtatsNiveaux As ETATS_NIVEAUX                                                         '�tats des niveaux
    
    EtatsEVEau As ETATS_EV_EAU                                                             '�tats de l'�lectro-vanne d'arriv�e d'eau
    
    API_ModeCouvercles As MODES_COUVERCLES                                  'retour API du mode des couvercles
    API_CycleCouvercles As CYCLES_COUVERCLES                                 'retour API du cycle des couvercles (cycle des ponts)
    ModeCouvercles As MODES_COUVERCLES                                         'mode des couvercles d�terminer par l'op�rateur
    'pour les �tats des couvercles voir les postes
    
    Temperatures As VarTemperatures                                                       'valeurs des temp�ratures

End Type

'***************************************************************************************************************************
'                                                                        ZONES DE LA  LIGNE
'***************************************************************************************************************************

'--- enregistrement du type des zones de la ligne ---
Public Type EnrZones
    NumZone As Integer                     'N� de la zone
    Codezone As String                      'Code de la zone
    LibelleZone As String                   'Libell� de la zone
    NumPremierPoste As Integer       'N� du premier poste
    NomPremierPoste As String         'Nom du premier poste
    NumDernierPoste As Integer        'N� du dernier poste
    NomDernierPoste As String          'Nom du dernier poste
    NbrPostes As Integer                    'Nombre de postes concern�s par la zone
End Type

'--- enregistrement du type des barres de la ligne ---
Public Type EnrBarres
    NumBarre As Integer                     'N� de la barre
    Libelle As String                   'Libell� de la barre
End Type

'***************************************************************************************************************************
'                                                                      GAMMES D'ANODISATION
'***************************************************************************************************************************

'--- enregistrement du type des d�tails des gammes d'anodisation ---
Public Type EnrDetailsGammesAnodisation
    
    NumLigne As Integer                                                  'n� de ligne
    NumZone As Integer                                                   'n� de zone
    
    TempsAuPosteSecondes As Long                              'temps au poste en secondes
    TempsAuPosteTexte As String                                    'temps au poste en texte au format HH:MM:SS
    
    TempsAlerteSecondes As Long                                   'temps d'alerte en secondes
    TempsAlerteTexte As String                                        'temps d'alerte en texte au format HH:MM:SS
    
    TempsEgouttageSecondes As Integer                        'temps d'�gouttage en secondes
    TempsEgouttageTexte As String                                 'temps d'�gouttage en texte au format MM:SS

    '********** UTILISER UNIQUEMENT EN PRODUCTION **********
    NumPosteReel As Integer                                           'n� de poste r�el utilis� dans la zone
                                                                                         '(cas des postes multiples)
    
    DecompteDuTempsAuPosteReelSecondes As String 'repr�sente la diff�rence entre le temps th�orique au poste
                                                                                         'et le temps r�el pass� dans le poste
                                                                                         'un nombre n�gatif apparait si la charge est rest� plus
                                                                                         'longtemps dans le poste que le temps th�orique pr�vu
                                                                                         'ATTENTION variable du type String volontairement
                                                                                         'Si "" alors il n'y a pas eu de temps de d�compter
    
    DecompteDuTempsAlerteReelSecondes As String     'repr�sente la diff�rence entre le temps th�orique d'alerte
                                                                                         'et le temps r�el pass� avant l'alerte
    
    FinDuTempsPosteReel As Boolean                            'TRUE = Indique la fin du temps au poste r�el
    DebutAlertePosteReel As Boolean                              'TRUE = Indique le d�but de l'alertes au poste r�el

End Type

'--- enregistrement du type 'matieres' ---
Public Type EnrMatieres
    Matiere As String                                                'mati�re
    TypeMatiere As String                                        'type de la mati�re
    CompositionMatiere As String                            'composition de la mati�re
    OrdrePourAffichage As Integer                           'ordre pour affichage
End Type

'--- enregistrement du type 'GammesAnodisation' (comprend �galement les d�tails) ---
Public Type EnrGammesAnodisation
    
    NumGamme As String                                              'n� de gamme
    DateCreationGamme As Date                                  'date de cr�ation de la gamme
    RefGamme As String                                                'r�f�rence de la gamme
    NomGamme As String                                              'nom de la gamme
    
    Designation As String                                               'd�signation de la gamme d'anodisation
    
    TMatieresGamme(1 To NBR_MATIERES_MAXI_PAR_GAMME) As String    'tableau contenant les mati�res de la gamme
    
    TempsAvantPostePrincipalTexte As String             'temps avant Anodisation en texte au format HH:MM:SS
    TempsPostePrincipalTexte As String                      'temps au poste d'anodisation en texte au format HH:MM:SS
    TempsApresPostePrincipalTexte As String             'temps apr�s Anodisation en texte au format HH:MM:SS
    TempsTotalPostesTexte As String                           'temps total des postes en texte au format HH:MM:SS
    TempsTotalEgouttagesTexte As String                    'temps total des �gouttages en texte au format HH:MM:SS
    TempsTotalGammeTexte As String                          'temps total de la gamme en texte au format HH:MM:SS
    
    TempsAvantPostePrincipalSecondes As Long        'temps avant Anodisation en secondes
    TempsPostePrincipalSecondes As Long                 'temps au poste d'anodisation en secondes
    TempsApresPostePrincipalSecondes As Long        'temps apr�s Anodisation en secondes
    TempsTotalPostesSecondes As Long                     'temps total des postes en secondes
    TempsTotalEgouttagesSecondes As Long              'temps total des �gouttages en secondes
    TempsTotalGammeSecondes As Long                    'temps total de la gamme en secondes
    
    '************************* PASSAGE DANS LES BAINS **************************

    PassageAnodisation As Boolean                           'indique un passage dans un des bains d'anodisation
    PassageSpectro As Boolean                                  'indique un passage dans le bain de spectrocoloration
    PassageOr As Boolean                                           'indique un passage dans le bain d'or
    PassageNoir As Boolean                                        'indique un passage dans le bain de noir
    
    '*************************** GAMME REDRESSEUR *****************************
    ModeUouI As MODES_U_OU_I                                 'mode de travail du redresseur U(tension)=0, I(intensit�)=1
    
    TDetailsPhases(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases
    
    '********** AVEC LES DETAILS POUR AVOIR LA GAMME COMPLETE **********
    TDetailsGammesAnodisation(1 To NBR_LIGNES_DETAILS_GAMMES_PRODUCTION) As EnrDetailsGammesAnodisation 'gamme d'anodisation � executer

    '********** UTILISER UNIQUEMENT EN PRODUCTION **********
    ChoixPosteAnodisation As CHOIX_POSTE_ANODISATION  'choix du poste d'anodisation (imposer au chargement)

End Type

'***************************************************************************************************************************
'                                                                                  CHARGES
'***************************************************************************************************************************

'--- d�tails des charges (ATTENTION � la correspondance avec le chargement) ---
Public Type DetailsCharges
    NumCommandeInterne As Long                                                        'n� de commande interne = GACLEUNIK
    NumGamme As Long                                                                  'n� de gamme provenant de Clipper
    Naf As Long                                                                       'N� affaire
    TypeReparation As String                                                                              'nombre de r�parations (champ texte volontaire)
    CodeClient As String                                                                                      'code du client
    NbrPieces As Double                                                                                        'nombre de pi�ces
    Designation As String                                                                                    'd�signation
    Matiere As String                                                                                           'mati�res des pi�ces
    Observations As String                                                                                 'observations
    NumLignesReferencesClient As String                                                         'n� de lignes des r�f�rences du client correspondant
                                                                                                                           'aux n� de lignes des travaux s�par�s par un tiret
End Type

'--- d�tails des fiches de production ---
Public Type DetailsFichesProduction
    
    NumPoste As Integer                                'num�ro du poste
    
    DateEntreePoste As Date                         'date d'entr�e dans le poste
    DateSortiePoste As Date                          'date de sortie du poste
    DateDebutEgouttage As Date                   'date de d�but de l'�gouttage
    DateFinEgouttage As Date                       'date de d�but de l'�gouttage
    
    TemperatureEnEntree As Single              'temp�rature en entr�e de bain (si temp�rature)
    TemperatureEnSortie As Single               'temp�rature en sortie de bain (si temp�rature)
    GrapheTemperature As String                  'graphe de la temp�ratrure
    
    URedresseur As Single                            'tension redresseur (si redresseur)
    IRedresseur As Single                              'intensit� redresseur (si redresseur)
    SensRedresseur As Integer                      'sens du redresseur en fonction du type de redresseur (voir l'�num�ration correspondante)
    GrapheRedresseur As String                    'graphe du redresseur
    
    AnalyseurEnEntree As Single                  'valeur de l'analyseur en entr�e de bain d'anodisation
    AnalyseurEnSortie As Single                   'valeur de l'analyseur en sortie de bain d'anodisation
    GrapheAnalyseur As String                      'graphe de l'analyseur
    
    AlarmesPoste As String                            'alarmes du poste concern�

End Type

'--- �tats des charges ---
Public Type etatsCharges

    DateEntreeEnLigne As Date                                      'date d'entr�e dans la ligne (g�n�ralement le chargement)
    DateArriveeAuDechargement As Date                       'date d'arriv�e au d�chargement
    
    NumBarre As Integer                                                 'num�ro de barre
    NumBarreInc As Integer                                                 'num�ro de barre incr�mental et journalier
    ChargePrioritaire As Boolean                                    'indique qu'il sagit  d'une charge prioritaire
                                                                                       'cette option est valid� au chargement
    
    DelaiSupStabilisationChargeSecondes As Integer   'd�lai suppl�mentaire de stabilisation de la charge en
                                                                                       'secondes en arr�t au poste pour �viter le mouvement
                                                                                       'pendulaire
    Options1 As Integer                                                   'cette valeur est transmise � l'automate et permet de g�rer
                                                                                       'les options 1 (vitesses de mont�e-descente, etc ...)
                                                                                       'pour certaines charges sur les ponts
    Options2 As Integer                                                   'cette valeur est transmise � l'automate et permet de g�rer
                                                                                       'les options 2 (vitesses de mont�e-descente, etc ...)
                                                                                       'pour certaines charges sur les ponts
   
    
    VitesseHaut As Integer
    VitesseBas As Integer
        
    
    TDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As DetailsCharges 'voir ci-dessus
    
    TGammesAnodisation As EnrGammesAnodisation  'gammes d'anodisation compl�te pour la production
    PtrZoneGammeAnodisation As Integer                     'pointeur de la zone de la gamme d'anodisation
    
    ModeUouI As MODES_U_OU_I                                 'mode de travail du redresseur U(tension)=0, I(intensit�)=1
    
    FinPhase4 As Boolean
    
    
    TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases  'pour renseigner le redresseur en entrant dans le bain
    
    TempsTotalGammeRedresseur As Long                   'temps total de la gamme redresseur en secondes
    
    NbrPostesTraites As Integer                                     'incr�mentation de 1 � chaque entr�e dans un poste
                                                                                       'sert d'index pour les d�tails des fiches de production
    TDetailsFichesProduction(1 To NBR_LIGNES_DETAILS_FICHES_PRODUCTION) As DetailsFichesProduction
    
    AlarmesLigne As String                                             'alarmes de la ligne (s�paration par -)
    
End Type

'***************************************************************************************************************************
'                                                                CHARGEMENT ET PREVISIONNEL
'***************************************************************************************************************************

'--- chargement ---
Public Type VarChargement
    TDetailsCharges(1 To NBR_LIGNES_DETAILS_CHARGES) As DetailsCharges                                                                                                  'voir ci-dessus
    TGammesAnodisation As EnrGammesAnodisation                                                                                                                                              'gammes d'anodisation compl�te pour la production
    TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1 To PHASES_GAMMES_REDRESSEURS.PH_T4) As DetailsPhases 'pour renseigner le redresseur en entrant dans le bain
End Type

'--- pr�visionnel ---
Public Type VarPrevisionnel
    
    ChoixIA As Integer                                                                'meilleur choix pour l'entr�e dans la ligne (moteur d'inf�rence)
    NumCommandeInterne As Long                                         'n� de commande interne
    NumGamme As Long
    Naf As Long
    TypeReparation As String                                                     'nombre de r�parations (champ texte volontaire)
    CodeClient As String                                                             'code du client
    NbrPieces As Double                                                               'nombre de pi�ces
    Designation As String                                                           'd�signation
    Observations As String                                                         'observations
    Matiere As String                                                                   'mati�re des pi�ces
    NumBarre As Integer                                                             'num�ro de barre
    
    NumGammeAnodisation As String                                        'n� de la gamme d'anodisation
    TGammesAnodisation As EnrGammesAnodisation              'gammes d'anodisation compl�te pour la production
    ChoixPosteAnodisation As CHOIX_POSTE_ANODISATION  'choix du poste d'anodisation

End Type

'--- ligne calcul du pr�visionnel ---
Public Type LignePrevisionnel
    
    NumLigne As Integer                                                             'n� de ligne utile pour le pr�visionnel
    NumGammeAnodisation As String                                        'n� de la gamme d'anodisation
    TempsPostePrincipalSecondes As Long                              'temps au poste principal en secondes
    
    PassageAnodisation As Boolean                                          'indique un passage dans un des bains d'anodisation
    PassageSpectro As Boolean                                                 'indique un passage dans le bain de spectrocoloration
    PassageOr As Boolean                                                          'indique un passage dans le bain d'or
    PassageNoir As Boolean                                                       'indique un passage dans le bain de noir

End Type

'***************************************************************************************************************************
'                                                               TRACABILITE DE LA PRODUCTION
'***************************************************************************************************************************

'--- enregistrement du type d�tails des charges de production ---
Public Type EnrDetailsChargesProduction
    
    NumCommandeInterne As Long                            'n� de commande interne
    TypeReparation As String                                        'nombre de r�parations (champ texte volontaire)
    DateEntreeEnLigne As Date                                    'date d'entr�e dans la ligne (g�n�ralement le chargement)
    DateArriveeAuDechargement As Date                     'date d'arriv�e au d�chargement
    NumBarre As Integer                                               'n� de barre
    NumLigne As Integer                                               'n� de ligne
    CodeClient As String                                                'code du client
    NbrPieces As Double                                                  'nombre de pi�ces
    Designation As String                                              'd�signation
    Matiere As String                                                     'mati�re des pi�ces
    NumLignesReferencesClient As String                   'n� de lignes des r�f�rences du client correspondant
                                                                                     'aux n� de lignes des travaux s�par�s par un tiret
    
    NumGammeAnodisation As String                          'n� de la gamme d'anodisation
    RefGammeAnodisation As String                            'r�f�rence de la gamme d'anodisation
    
    NumFicheProduction As String                                'n� de la fiche de production
    ChargePrioritaire As Boolean                                  'indique qu'il sagit  d'une charge prioritaire
    AlarmesLigne As String                                           'alarmes de la ligne
    ControleColmatage As Integer                                 'contr�le du colmatage (valeur de 0 � 5)
    ControleEpaisseurAnodisation As Integer               'contr�le de l'�paisseur d'anodisation (valeur de 0 � 100)
    ControleColoration As String                                    'contr�le de la coloration (20 caract�res)
    ControleObservations As String                               'observations sur les contr�les (50 caract�res)

End Type

'--- enregistrement du type d�tails des gammes de production ---
Public Type EnrDetailsGammesProduction
    NumFicheProduction As String                                     'n� de la fiche de production
    NumLigne As Integer                                                    'n� de ligne
    NumZone As Integer                                                     'n� de la zone
    TempsAuPosteTexte As String                                     'Temps au poste en texte au format HH:MM:SS
    TempsEgouttageTexte As String                                   'Temps d'�gouttage en texte au forma MM:SS
    TempsAuPosteSecondes As Long                                'Temps au poste en secondes
    TempsEgouttageSecondes As Integer                          'Temps d'�gouttage en secondes
    DecompteDuTempsAuPosteReelSecondes As String  'Repr�sente la diff�rence entre le temps th�orique au poste
                                                                                          'et le temps r�el pass� dans le poste
                                                                                          'un nombre n�gatif apparait si la charge est rest� plus
                                                                                          'longtemps dans le poste que le temps th�orique pr�vu
                                                                                          'ATTENTION variable du type String volontairement
                                                                                          'Si "" alors il n'y a pas eu de temps de d�compter
    NumPosteReel As Integer                                            'N� de poste r�el utilis� dans la zone (cas des postes multiples)
End Type

'--- enregistrement du type d�tails des fiches de production ---
Public Type EnrDetailsFichesProduction
    NumFicheProduction As String                  'n� de la fiche de production
    NumLigne As Integer                                 'n� de ligne
    NumPoste As Integer                                 'n� du poste
    DateEntreePoste As Date                          'date d'entr�e dans le poste
    DateSortiePoste As Date                           'date de sortie du poste
    DateDebutEgouttage As Date                    'date de d�but de l'�gouttage
    DateFinEgouttage As Date                        'date de d�but de l'�gouttage
    TemperatureEnEntree As Single               'temp�rature en entr�e de bain (si temp�rature)
    TemperatureEnSortie As Single                'temp�rature en sortie de bain (si temp�rature)
    GrapheTemperature As String                   'graphe de la temp�ratrure
    URedresseur As Single                             'tension redresseur (si redresseur)
    IRedresseur As Single                              'intensit� redresseur (si redresseur)
    SensRedresseur As Integer                      'sens du redresseur en fonction du type de redresseur (voir l'�num�ration correspondante)
    GrapheRedresseur As String                    'graphe du redresseur
    AnalyseurEnEntree As Single                   'valeur de l'analyseur en entr�e de bain d'anodisation
    AnalyseurEnSortie As Single                    'valeur de l'analyseur en sortie de bain d'anodisation
    GrapheAnalyseur As String                       'graphe de l'analyseur
    AlarmesPoste As String                             'alarmes du poste concern�
End Type

'--- enregistrement du type d�tails des phases de production ---
Public Type EnrDetailsPhasesProduction
    NumFicheProduction As String                  'n� de la fiche de production
    NumRedresseur As Integer                       'n� du redresseur
    ModeUouI As MODES_U_OU_I                 'mode tension ou intensit�
    NumPhase As Integer                                'num�ro de la phase
    TempsPhase As Integer                            'temps de la phase
    UPhase As Single                                      'U de la phase
    IPhase As Single                                        'I de la phase
End Type

'***************************************************************************************************************************
'                                                                                  DEFAUTS
'***************************************************************************************************************************

'--- enregistrement du type des d�fauts ---
Public Type EnrDefauts
    
    NumDefaut As Integer                                          'n� du d�faut
    
    SignalerOuiNon As Boolean                                 'indique si le d�faut doit �tre ou pas signaler
    
    GyrophareOuiNon As Boolean                              'indique que le d�faut d�clenche ou non le gyrophare de la ligne
    
    KlaxonOuiNon As Boolean                                    'indique que le d�faut d�clenche ou non le klaxon de la ligne
    
    MessageVocalOuiNon As Boolean                       'indique si le d�faut d�clenche ou non l'envoi d'un message vocal
    
    AfficheurOuiNon As Boolean                                'indique si le d�faut concerne l'afficheur � leds rouge
    
    InformationsAPI As String                                     'informations API concernant le d�faut
    LibelleDefaut As String                                         'libell� du d�faut
    LibelleDefautAfficheur As String                           'libell� du d�faut destin� � l'afficheur

    TNumIntervenants(1 To 5) As Integer                   'tableau contenant les num�ros (index de personne) des intervenants
    
    '********** NON MEMORISER DANS LA BASE DE DONNEES, UTILISER UNIQUEMENT EN INTERNE **********
    AntiRebondGyrophare As Boolean                      'anti-rebond de signalisation du d�faut sur gyrophare
    AntiRebondKlaxon As Boolean                            'anti-rebond de signalisation du d�faut sur klaxon
    AntiRebondTra�abiliteAlarmes As Boolean         'anti-rebond de signalisation du d�faut dans la table de tra�abilit�
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

'--- variable des commandes op�rateur ---
Public Type VarCommandesOperateur
    TypeCycle As TYPES_CYCLES                           'type de cycle fonction de l'�num�ration TYPES_CYCLES
    NumPont As Integer                                             'num�ro du pont concern� par la commande
    NumPosteDepart As Integer                                'num�ro du poste de d�part concern� par la commande
    NumPosteArrivee As Integer                               'num�ro du poste d'arriv�e concern� par la commande
    TempsEgouttageSecondes As Integer                'temps d'�gouttage en secondes concern� par la commande
End Type

'***************************************************************************************************************************
'                                                                     MOTEUR D'INFERENCE
'***************************************************************************************************************************

'--- variable de l'ordre de sortie des charges ---
Public Type VarOrdreSortieCharges
    
    NumPoste As Integer                                                                      'num�ro du poste
    NumCharge As Integer                                                                    'num�ro de charge au poste
    
    NumPosteArrivee As Integer                                                          'num�ro du poste d'arriv�e pr�vu
    NumChargePosteArrivee As Integer                                              'num�ro de charge au poste d'arriv�e
    
    DecompteDuTempsAuPosteReelSecondes As String                    'd�compte du temps au poste r�el en secondes
    Condamnation As Boolean                                                             'TRUE=Poste condamn�
                                                                                                            'FALSE=Poste en fonctionnement normal
    NumPont As Integer                                                                        'num�ro du pont th�orique de la pr�misse choisi
                                                                                                            'pour le futur mouvement lorsque le temps sera � 0
                                                                                                    
End Type

'--- variable des informations sur les postes d'anodisation ---
Public Type VarInformationsPostesAnodisation
    NumCharge As Integer                                                                    'num�ro de charge au poste
    DecompteDuTempsAuPosteReelSecondes As String                    'd�compte du temps au poste r�el en secondes
    Condamnation As Boolean                                                             'TRUE=Poste condamn�
                                                                                                            'FALSE=Poste en fonctionnement normal
End Type

'--- variable du moteur d'inf�rence ---
Public Type VarMoteurInference
    
    '--- chargement ---
    ProchainNumPosteChargementSiAnodisationC13Impose As Integer     'prochain n� de poste au chargement si le poste
                                                                                                                     'd'anodisation C13 est impos� dans la gamme
    ProchainNumPosteChargementSiAnodisationC14Impose As Integer     'prochain n� de poste au chargement si le poste
                                                                                                                     'd'anodisation C14 est impos� dans la gamme
    ProchainNumPosteChargementSiAnodisationC15Impose As Integer     'prochain n� de poste au chargement si le poste
                                                                                                                     'd'anodisation C15 est impos� dans la gamme
    ProchainNumPosteChargementSiAnodisationC16Impose As Integer     'prochain n� de poste au chargement si le poste
                                                                                                                     'd'anodisation C16 est impos� dans la gamme
    
    ProchainNumPosteChargementSiAnodisationAutomatique As Integer   'prochain n� de poste au chargement si le poste
                                                                                                                     'd'anodisation est automatique dans la gamme
    ProchainNumPosteChargement As Integer                                             'indique le prochain num�ro de poste ou se fera
                                                                                                                     'le chargement
    
    '--- ordre de sortie des charges dans la ligne ---
    TOrdreSortieCharges(1 To CHARGES.C_NUM_MAXI) As VarOrdreSortieCharges
                                                                                                                    'tableau contenant l'ordre de sortie des charges
                                                                                                                    'ce tableau est tri� directement du temps le plus
                                                                                                                    'court au temps le plus long
                                                                                                                    'CHARGES.C_NUM_MAXI correspondant au
                                                                                                                    'nombre de charges maxi dans la ligne
    '--- ordre de sortie pour les ponts ---
    TOrdreSortiePonts(PONTS.P_1 To PONTS.P_2, 1 To CHARGES.C_NUM_MAXI) As VarOrdreSortieCharges
                                                                                                                    'tableau contenant l'ordre de sortie des charges
                                                                                                                    'ce tableau est tri� directement du temps le plus
                                                                                                                    'court au temps le plus long
                                                                                                                    'CHARGES.C_NUM_MAXI correspondant au
                                                                                                                    'nombre de charges maxi dans la ligne
    
    '--- informations compl�tes sur les postes d'anodisation ---
    'sert � connaitre le moment opportum pour l'entr�e d'une charge dans la ligne
    TInformationsPostesAnodisation(POSTES.P_C13 To POSTES.P_C16) As VarInformationsPostesAnodisation
                                                                                                                   'tableau contenant les informations sur les
                                                                                                                   'postes d'anodisation
End Type

'***************************************************************************************************************************
'                                                                  TRACABILITE DES ALARMES
'***************************************************************************************************************************

'--- enregistrement du type de la tra�abilit� des alarmes ---
Public Type EnrTra�abiliteAlarmes
    ClePrimaire As Long                         'cl� primaire
    NumDefaut As Integer                       'n� du d�faut
    DateDetectionDefaut As Date           'date de d�tection du d�faut
End Type

'***************************************************************************************************************************
'                                                                        JOURNEES TYPES
'***************************************************************************************************************************

'--- variable d'un cycle de 24 heures pour les journ�es types ---
Public Type VarCycle24HJourneesTypes
    TypeDeJournee As JOURNEES_TYPES                                               'type de journ�e par cuve
    TTopsDebutPompe(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de d�but d'un cycle pompe
    TTopsFinPompe(1 To NBR_TOPS_POSSIBLES) As String * 14           'tops de fin d'un cycle pompe
    TCyclesPompe(1 To NBR_TOPS_POSSIBLES) As Integer                   'cycles de la pompe
    TTopsDebutChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14  'tops de d�but d'un cycle chauffage
    TTopsFinChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de fin d'un cycle chauffage
    TModesChauffage(1 To NBR_TOPS_POSSIBLES) As Integer               'modes du chauffage
End Type

'***************************************************************************************************************************
'                                                                PROGRAMMATEUR CYCLIQUE
'***************************************************************************************************************************

'--- variable d'un cycle de 24 heures pour le programmateur cyclique ---
Public Type VarCycle24HProgCyclique
    TypeDeJournee As JOURNEES_TYPES                                               'type de journ�e par cuve
    TTopsDebutPompe(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de d�but d'un cycle pompe
    TTopsFinPompe(1 To NBR_TOPS_POSSIBLES) As String * 14            'tops de fin d'un cycle pompe
    TCyclesPompe(1 To NBR_TOPS_POSSIBLES) As Integer                    'cycles de la pompe
    TTopsDebutChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14   'tops de d�but d'un cycle chauffage
    TTopsFinChauffage(1 To NBR_TOPS_POSSIBLES) As String * 14       'tops de fin d'un cycle chauffage
    TModesChauffage(1 To NBR_TOPS_POSSIBLES) As Integer               'modes du chauffage
End Type

'***************************************************************************************************************************
'                                                                               ANNEXES
'***************************************************************************************************************************

'--- num�ros des d�fauts ---
'Public Type NumDefautsAnnexes
    
'End Type

'--- �tats des annexes ---
Public Type EtatsAnnexes
    
    'TNumDefauts As numdefautsAnnexes
    
    API_ChangementsEVBrillantage As Boolean                                     'TRUE=Indique un changement (hors �tats)
    API_EtatsEVBrillantage As String * 16                                                'retour en binaire du mot de l'API contenant les �tats
    EtatsEVBrillantage As ETATS_EV_BRILLANTAGE                              'fonction de l'�num�ration
    PeriodiciteEVBrillantage As Integer                                                    'p�riodicit� de mise en marche de l'�lectro-vanne d'air dans le bain de brillantage
    TempsMarcheEVBrillantage As Integer                                              'temps de marche de l'�lectro-vanne d'air dans le bain de brillantage
    ModeEVBrillantage As MODES_EV_BRILLANTAGE                           'fonction de l'�num�ration

    API_ChangementsEVEauLigne As Boolean                                        'TRUE=Indique un changement (hors �tats)
    EtatsEVEauLigne As ETATS_EV_EAU_LIGNE                                     'fonction de l'�num�ration
    ModeEVEauLigne As MODES_EV_EAU_LIGNE                                  'fonction de l'�num�ration
    
    API_ChangementsCompresseurP1 As Boolean                                  'TRUE=Indique un changement (hors �tats)
    EtatsCompresseurP1 As ETATS_COMPRESSEURS_PONTS              'fonction de l'�num�ration
    ModeCompresseurP1 As MODES_COMPRESSEURS_PONTS           'fonction de l'�num�ration
    
    API_ChangementsEclairageP1 As Boolean                                         'TRUE=Indique un changement (hors �tats)
    EtatsEclairageP1 As ETATS_ECLAIRAGE_PONTS                               'fonction de l'�num�ration
    ModeEclairageP1 As MODES_ECLAIRAGE_PONTS                            'fonction de l'�num�ration
    
    API_ChangementsCompresseurP2 As Boolean                                  'TRUE=Indique un changement (hors �tats)
    EtatsCompresseurP2 As ETATS_COMPRESSEURS_PONTS              'fonction de l'�num�ration
    ModeCompresseurP2 As MODES_COMPRESSEURS_PONTS           'fonction de l'�num�ration
    
    API_ChangementsEclairageP2 As Boolean                                         'TRUE=Indique un changement (hors �tats)
    EtatsEclairageP2 As ETATS_ECLAIRAGE_PONTS                               'fonction de l'�num�ration
    ModeEclairageP2 As MODES_ECLAIRAGE_PONTS                             'fonction de l'�num�ration
    
    TAPI_ChangementsNiveauxRetentions(NIVEAUX_RETENTIONS.NR_STOCKAGE_STATION To NIVEAUX_RETENTIONS.NR_LAVEUR) As Boolean                                          'TRUE=Indique un changement (hors �tats)
    TEtatsNiveauxRetentions(NIVEAUX_RETENTIONS.NR_STOCKAGE_STATION To NIVEAUX_RETENTIONS.NR_LAVEUR) As ETATS_NIVEAUX_RETENTIONS                          'fonction de l'�num�ration

End Type

'***************************************************************************************************************************
'                                                                   COMMANDES INTERNE
'***************************************************************************************************************************

'--- enregistrement du type "CommandesInterne" ---
Public Type EnrCommandesInterne
    
    NumCommandeInterne As Long             'n� de commande interne
    CodeClient As String                                'Code du client
    Designation As String                               'd�signation
    NbrPieces As Double                                   'nombre de pi�ces
    Matiere As String                                      'mati�re des pi�ces

End Type

'***************************************************************************************************************************
'                                                                         PHASES DE CLIPPER
'***************************************************************************************************************************

'--- enregistrement du type "PhasesClipper" ---
Public Type EnrPhasesClipper
    GaCLeUnik  As Long                          'cl� unique pour rechercher une affaire
    CoFrais As String                               'centre de frais
    CoCli As String                                   'code client
    NomClient As String                           'nom du client
    Piece As String                                  'r�f�rence de la pi�ce
    QteAf As Double                                    'quantit� de pi�ces
    Desa1 As String                                 'd�signation de la pi�ce
    DateLance As String                          'date de lancement format texte JJ/MM/AAAA
    Matiere As String                               'mati�re de la pi�ce
    GamObs As String                             'observations
    NumGamme As Long
    Naf As Long
End Type

'***************************************************************************************************************************
'                                                               GRAPHES DE PRODUCTION
'***************************************************************************************************************************

'--- tra�abilit� ---
Public Type Tra�abilite
    DateDuPoint As Date                                  'date du point (construction de l'axe des x)
    NumPhase As Byte                                     'num�ro de la phase
    EtatRedresseur As Byte                             '�tat du redresseur (0=Arr�t, 1=Pause, 2= Marche, 3=D�faut)
    Tension As Byte                                          'tension mesur�e (valeur en d�cimale fois 10)
    Intensite As Integer                                     'intensit� mesur�e
    Temperature As Integer                               'temp�rature mesur�e (valeur en d�cimale fois 10)
End Type

'--- renseignements d'un graphe ---
Public Type RenseignementsGraphe
    NumFicheProduction As String                    'n� de la fiche de production
    DateEntreeEnLigne As Date                        'date d'entr�e dans la ligne (g�n�ralement le chargement)
    NumRedresseur As Integer                         'num�ro du redresseur
End Type

'***************************************************************************************************************************
'                                                                           ETUVE en C33/C34
'***************************************************************************************************************************

'--- entr�es API de l'�tuve C33/C34 ---
Public Type EntreesEtuveC33C34_OLD
    
    DesequilibrePhases As Boolean                             'd�s�quilibre des phases

    SurchauffeElementChauffant1 As Boolean              'surchauffe de l'�l�ment chauffant 1
    SurchauffeElementChauffant2 As Boolean              'surchauffe de l'�l�ment chauffant 2
    SurchauffeElementChauffant3 As Boolean              'surchauffe de l'�l�ment chauffant 3
    
    DefautMarcheVentilateur As Boolean                      'd�faut de marche du ventilateur

End Type

'--- sorties API de l'�tuve C33/C34 ---
'Public Type SortiesEtuveC33C34

'End Type

'--- num�ros des d�fauts ---
Public Type NumDefautsEtuveC33C344_OLD
    
    NumDefautDesequilibrePhases As Integer

    NumDefautSurchauffeElementChauffant1 As Integer
    NumDefautSurchauffeElementChauffant2 As Integer
    NumDefautSurchauffeElementChauffant3 As Integer
    
    NumDefautDefautMarcheVentilateur As Integer
    
End Type

Public Type EtatsEtuveC33C344_OLD
    
    TNumDefauts As NumDefautsEtuveC33C344_OLD                             'num�ros des d�fauts
    
    ManuelAutomatique As Boolean                                                'manuel / automatique de l'�tuve
    
    DepartCycle As Boolean                                                            'd�part cycle
    CycleEnCours As Boolean                                                         'cycle en cours
    CycleTermine As Boolean                                                         'cycle termin�
    
End Type

'***************************************************************************************************************************
'                                                                                DIVERS
'***************************************************************************************************************************

'--- manipulations dans la fen�tre gestion de la r�gulation ---
Public Type ManipsGestionRegulation
    AppareillageConcerne As Boolean           'FALSE=pompe, TRUE=chauffage
    CyclesPompe As Integer
    ModesChauffage As Integer
End Type

'--- manipulations dans la fen�tre du programmateur cyclique ---
Public Type ManipsProgCyclique
    AppareillageConcerne As Boolean           'FALSE=pompe, TRUE=chauffage
    CyclesPompe As Integer
    ModesChauffage As Integer
End Type

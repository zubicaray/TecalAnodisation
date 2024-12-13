Attribute VB_Name = "MPVariables"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE DES VARIABLES PUBLIQUES
' Nom                    : MPVariables.bas
' Date de création : 31/07/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'*** VARIABLES ***

Public CORRESPONDANCES_IDX_AUTOMATE(11) As Integer
Public CORRESPONDANCES_IDX_CUVES_API(18) As Integer
'--- type de PC ---
Public TypePC As TYPES_PC                                       'indique le type de PC première donnée du fichier de configuration

'--- type de BD (base de données) ---
Public TypeBD As TYPES_BD                                          'indique le type de bases de données



'--- chemins ---
Public RepFicAnodisation  As String                            'répertoire des fichiers de l'anodisation
Public RepImagesAnodisation As String                      'répertoire des images de l'anodisation
Public RepGraphesProductionLocal As String             'répertoire contenant les graphes de production de l'anodisation sur le PC en local
Public RepGraphesProductionServeur As String         'répertoire contenant les graphes de production de l'anodisation sur le serveur
Public RepFicClipper As String                                     'répertoire des fichiers de Clipper
Public CONFIG_FILE As String

'--- divers pour construction du programme ---
Public Bidon As Variant
Public varConfig As Variant
'--- renseignements sur l'ordinateur et son utilisateur ---
Public NOM_ORDINATEUR As String                          'nom de l'ordinateur
Public NOM_UTILISATEUR As String                          'nom de l'utilisateur de l'ordinateur
Public NOM_ORDINATEUR_UTILISATEUR As String  'nom de l'ordinateur et utilisateur séparé par le symbole |

'--- heure et date du système pour l'affichage de l'heure en bas de l'écran ---
Public HeureSysteme As String               'heure système en chaine "HH:MM:SS"
Public MemHeureSysteme As String       'mémoire de l'heure système
Public DateSysteme As String                 'date système en chaine "Jour XX Mois Année"
Public MemDateSysteme As String         'mémoire de la date système

'--- heure et date pour le noyau central ---
Public Maintenant As Currency                'référenciel de temps numérique AAAAMMJJHHMMSS
Public DateMaintenant As String * 10      'référenciel de la date en chaine "JJ/MM/AAAA"
Public HeureMaintenant As String * 8      'référenciel de l'heure en chaine "HH:MM:SS"
Public AnneesMaintenant As Integer       'référenciel des années en numérique
Public MoisMaintenant As Integer           'référenciel des mois en numérique
Public JoursMaintenant As Integer           'référenciel des jours en numérique
Public HeuresMaintenant As Integer        'référenciel des heures en numérique
Public MinutesMaintenant As Integer       'référenciel des minutes en numérique
Public SecondesMaintenant As Integer    'référenciel des secondes en numérique

'--- pour connaitre les fenêtres réellement chargées ---
Public FProgCycliqueChargee As Boolean
Public MemDateProgCyclique As String * 10         'mémoire de la date pour changer le prog. cyclique

'--- noyau central ---
Public PremierPassageNoyauCentral As Boolean 'indique le premier passage dans le noyau central

'--- affichage complet des outils de la Fenetre principale ---
Public AffichageCompletOutilsFPrincipale As Boolean 'indique un affichage complet des outils de la fenêtre principale

'--- pour la configuration ---
Public MemMenuPrincipalNavigateur As Integer   'mémoire de la position du menu principal du navigateur
Public MemSousMenuNavigateur As Integer         'mémoire de la position du sous menu du navigateur

'--- pour l'entretien des fichiers des graphes de production ---
Public EntretienGraphesProduction As Boolean

'--- pour les paramètres du logiciel (clé "Configuration" dans la base des registres) ---
Public RepLocalBD As String                              'répertoire en local des bases de données
Public RepDistantBD As String                           'répertoire en distant des bases de données
Public ModeDeConnexion As Integer                  '0=en réseau, 1=en autonome

Public PARAMETRES_CONNEXION_BD_ANODISATION_SQL As String
Public PARAMETRES_CONNEXION_BD_CLIPPER_HF As String
Public MODE_SECOURS As Boolean

Public SuppressionMotsDePasse As Boolean   'suppression des demandes de mots de passe dans le logiciel
Public MotDePasseDirection As String               'mot de passe de la direction
Public MotDePassePersonnel As String             'mot de passe personnel
Public UniteMonetaire As Integer                        '0=Francs français, 1=Euros
Public IndicePrestationParDefaut As Integer       'indice de la prestation par défaut (0=CHROMAGE, etc...)
Public LibellePrestationParDefaut As String       'libellé de la prestation par défaut ("CHROMAGE", etc...)
Public NbrLignesMaxiAExtraire As Long             'nombre de lignes maxi à extraire pour les grandes tables ou requêtes

Public DISTANCE_SECURITE As Long                  ' distance de sécurité pour l'anti-collision
Public DEBUG_MODE  As Boolean

'--- pour les mots de passe ---
Public TypeDeMotDePasse As Boolean
    
'--- entrée automatique des charges ---
Public EntreeAutomatiqueCharges As Boolean  'entrée automatique des charges
    
'--- pour le copier / coller spécial ---
Public NumFenetreEnCopie As Long                  'indique le numéro de fenetre en copie pour le copier / coller spécial
Public CleDeCopie As Variant                             'indique la clé de copie (N° de commande interne, N° du devis, ...)

'--- pour l'impression des états ---
Public OptionImpressionChoisie As Integer       'option d'impression pour les états
Public MargeGaucheTwips As Long                   'marge gauche en twips pour l'impression
Public MargeHauteTwips As Long                      'marge haute en twips pour l'impression
Public MargeDroiteTwips As Long                      'marge droite en twips pour l'impression
Public MargeBasseTwips As Long                     'marge basse en twips pour l'impression
Public PersonneEmettrice As String                   'nom de la personne émettrice

'--- pour les modifications avant impression ---
Public MemReperefenetreCritereRecherche As String 'mémoire du repère de la fenetre d'appel et du critère de recherche

'--- caractères spéciaux ---
Public CARACTERE_PHI As String * 1                   'caractère phi (diamètre)
Public CARACTERE_FRANC As String * 1              'caractère pour le franc
Public CARACTERE_EURO As String * 1               'caractère pour l'euro

'--- pour le filtrage des touches ---
Public ModeSurFrappe As Boolean

'--- limites du tableau des zones ---
Public LIMITE_BASSE_ZONES As Integer             'limite basse du tableau des zones
Public LIMITE_HAUTE_ZONES As Integer             'limite haute du tableau des zones

'--- limites du tableau des barres ---
Public LIMITE_BASSE_BARRES As Integer             'limite basse du tableau des barres
Public LIMITE_HAUTE_BARRES As Integer             'limite haute du tableau des barres
'--- images ---
Public ImgFondDeFenetre As Picture                     'image de fond standard d'un fenêtre
Public ImgFondDeFenetreXP As Picture                 'image de fond standard d'une fenêtre type XP
Public ImgFondEspace As Picture                          'image de fond de l'espace

Public ImgFondOrange1 As Picture                       'image de fond en orange 1
Public ImgFondOrange2 As Picture                       'image de fond en orange 2

Public ImgFondBleu1 As Picture                           'image de fond en bleu 1
Public ImgFondBleu2 As Picture                           'image de fond en bleu 2

Public ImgFondVert1 As Picture                             'image de fond en vert 1
Public ImgFondVert2 As Picture                             'image de fond en vert 2

Public ImgFondGris1 As Picture                             'image de fond en gris 1
Public ImgFondGris2 As Picture                             'image de fond en gris 2

Public ImgFondDesBoutons As Picture                 'image de fond standard pour les boutons

'--- pour le report des défauts sur le gyrophare et le klaxon ---
Public SignalerDefautSurGyrophare As Boolean   'TRUE = Indique une demande de déclenchement du gyrophare
Public SignalerDefautSurKlaxon As Boolean         'TRUE = Indique une demande de déclenchement du klaxon

'--- pour la gestion de la fin de journée ---
Public GestionFinDeJourneeEnCours As Boolean 'TRUE = Indique une demande de fin de journée
                                                                                 'FALSE = Gestion normale de la ligne

'--- alarmes de la ligne en cours
'    cette variable contient toutes les alarmes (séparées par le séparateur des alarmes) de la ligne à l'instant x
Public AlarmesLigneEnCours As String

'--- priorité de l'afficheur pour les alertes ---
Public PrioriteAfficheurPourAlertes As Boolean       'FALSE = pas de priorité d'affichage des alertes donc affichage des défauts
                                                                                  'TRUE = priorité d'affichage des alertes

'--- temps de compensation d'anodisation en minutes ---
Public TempsCompensationAnodisationMinutes As Integer

'--- mode d'affichage du synoptique et du chargement prévisionnel ---
'permet l'affichage entre le numéro de barre et de charge
Public ModeAffichageSynoptique As MODES_AFFICHAGES_SYNOPTIQUE

'*** TABLEAUX ***

'--- images des éléments ---
Public TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE To ECHELLES_24H.E_VENTILATION_CHAUFFAGE) As Picture 'images des échelles 24 heures pour le programmateur cyclique
Public TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_VERT To IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_EXCLUS) As Picture
Public TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_VERT To IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_EXCLUS) As Picture

'--- mémoires temporaires des enregistrements des tables pour lecture ou écriture de données ---
Public TTempEnrGammesAnodisation As EnrGammesAnodisation
Public TTempEnrDetailsGammesAnodisation() As EnrDetailsGammesAnodisation
Public TTempEnrCommandesInterne As EnrCommandesInterne
Public TTempEnrDetailsChargesProduction() As EnrDetailsChargesProduction
Public TTempEnrDetailsGammesProduction() As EnrDetailsGammesProduction
Public TTempEnrDetailsPhasesProduction() As EnrDetailsPhasesProduction
Public TTempEnrDetailsFichesProduction() As EnrDetailsFichesProduction
Public TTempEnrPhasesClipper As EnrPhasesClipper

'--- pour l'impression des états ---
Public TCriteresImpression(1 To 10) As Variant      'critères d'impression (paramètres)

'--- matières ---
Public TMatieres(1 To 50) As EnrMatieres

'--- zones ---
Public TZones() As EnrZones                                 'zones de la ligne d'anodisation
'--- barres ---
Public TBarres() As EnrBarres                                 'zones de la ligne d'anodisation

'--- actions ---
Public TActions(NUM_ACTION_NOP To NUM_ACTION_FCY) As EnrActions

'--- image de la mémoire des cycles des ponts (sans adaptation pour les paramètres) ---
Public TImageAPICyclesPonts(PONTS.P_1 To PONTS.P_2, 1 To NBR_LIGNES_CYCLES_PONTS) As Integer

'--- états des ponts ---
Public TEtatsPonts(PONTS.P_1 To PONTS.P_2) As EtatsPonts

'--- caractéristiques des cuves ---
Public TCaracteristiquesCuves(CUVES.C_C00 To CUVES.C_C02) As CaracteristiquesCuves

'--- états de la ligne ---
Public TEtatsLigne As EtatsLigne                          'états de la ligne

'--- états des cuves ---
Public TEtatsCuves(CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As EtatsCuves

'--- états des postes ---
Public TEtatsPostes(POSTES.P_CHGT_1 To DERNIER_POSTE) As EtatsPostes

'--- états des charges ---
Public TEtatsCharges(CHARGES.C_NUM_MINI To CHARGES.C_NUM_MAXI) As etatsCharges

'--- prémisses ---
Public TPremisses(POSTES.P_CHGT_1 To DERNIER_POSTE, POSTES.P_CHGT_1 To DERNIER_POSTE) As VarPremisses

'--- chargement ---
Public TChargement As VarChargement

'--- prévisionnel ---
Public TPrevisionnel(1 To NBR_LIGNES_PREVISIONNEL) As VarPrevisionnel

'--- états des redresseurs ---
Public TEtatsRedresseurs(REDRESSEURS.R_C13 To REDRESSEURS.R_C19) As EtatsRedresseurs

'--- états des annexes (ventilation, volet de compensation, ...) ---
Public TEtatsAnnexes As EtatsAnnexes

'--- défauts ---
Public TDefauts(DEFAUTS.NUM_MINI To DEFAUTS.NUM_MAXI) As EnrDefauts

'--- commandes opérateur ---
Public TCommandesOperateur(1 To 20) As VarCommandesOperateur  'commandes de l'opérateur

'--- moteur d'inférence ---
Public TMoteurInference As VarMoteurInference      'moteur d'inférence (contient toutes les données)


'--- journées types ---
Public TJourneesTypes(CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION, JOURNEES_TYPES.J_ARRET To JOURNEES_TYPES.J_REPRISE) As VarCycle24HJourneesTypes

'--- programmateur cyclique ---
Public TProgCyclique(1 To NBR_JOURS_PROG_CYCLIQUE, CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As VarCycle24HProgCyclique

'--- mémorisation de divers manipulations ---
Public VManipsGestionRegulation As ManipsGestionRegulation
Public VManipsProgCyclique As ManipsProgCyclique

'--- renseignements sur le graphe à imprimer ---
Public TRenseignementsGraphe As RenseignementsGraphe

'--- tableau des canaux des fichiers de traçabilité ---
Public TCanauxFichiersTraçabilite(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Integer






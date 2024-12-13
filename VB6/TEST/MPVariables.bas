Attribute VB_Name = "MPVariables"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE DES VARIABLES PUBLIQUES
' Nom                    : MPVariables.bas
' Date de cr�ation : 31/07/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'*** VARIABLES ***

Public CORRESPONDANCES_IDX_AUTOMATE(11) As Integer
Public CORRESPONDANCES_IDX_CUVES_API(18) As Integer
'--- type de PC ---
Public TypePC As TYPES_PC                                       'indique le type de PC premi�re donn�e du fichier de configuration

'--- type de BD (base de donn�es) ---
Public TypeBD As TYPES_BD                                          'indique le type de bases de donn�es



'--- chemins ---
Public RepFicAnodisation  As String                            'r�pertoire des fichiers de l'anodisation
Public RepImagesAnodisation As String                      'r�pertoire des images de l'anodisation
Public RepGraphesProductionLocal As String             'r�pertoire contenant les graphes de production de l'anodisation sur le PC en local
Public RepGraphesProductionServeur As String         'r�pertoire contenant les graphes de production de l'anodisation sur le serveur
Public RepFicClipper As String                                     'r�pertoire des fichiers de Clipper
Public CONFIG_FILE As String

'--- divers pour construction du programme ---
Public Bidon As Variant
Public varConfig As Variant
'--- renseignements sur l'ordinateur et son utilisateur ---
Public NOM_ORDINATEUR As String                          'nom de l'ordinateur
Public NOM_UTILISATEUR As String                          'nom de l'utilisateur de l'ordinateur
Public NOM_ORDINATEUR_UTILISATEUR As String  'nom de l'ordinateur et utilisateur s�par� par le symbole |

'--- heure et date du syst�me pour l'affichage de l'heure en bas de l'�cran ---
Public HeureSysteme As String               'heure syst�me en chaine "HH:MM:SS"
Public MemHeureSysteme As String       'm�moire de l'heure syst�me
Public DateSysteme As String                 'date syst�me en chaine "Jour XX Mois Ann�e"
Public MemDateSysteme As String         'm�moire de la date syst�me

'--- heure et date pour le noyau central ---
Public Maintenant As Currency                'r�f�renciel de temps num�rique AAAAMMJJHHMMSS
Public DateMaintenant As String * 10      'r�f�renciel de la date en chaine "JJ/MM/AAAA"
Public HeureMaintenant As String * 8      'r�f�renciel de l'heure en chaine "HH:MM:SS"
Public AnneesMaintenant As Integer       'r�f�renciel des ann�es en num�rique
Public MoisMaintenant As Integer           'r�f�renciel des mois en num�rique
Public JoursMaintenant As Integer           'r�f�renciel des jours en num�rique
Public HeuresMaintenant As Integer        'r�f�renciel des heures en num�rique
Public MinutesMaintenant As Integer       'r�f�renciel des minutes en num�rique
Public SecondesMaintenant As Integer    'r�f�renciel des secondes en num�rique

'--- pour connaitre les fen�tres r�ellement charg�es ---
Public FProgCycliqueChargee As Boolean
Public MemDateProgCyclique As String * 10         'm�moire de la date pour changer le prog. cyclique

'--- noyau central ---
Public PremierPassageNoyauCentral As Boolean 'indique le premier passage dans le noyau central

'--- affichage complet des outils de la Fenetre principale ---
Public AffichageCompletOutilsFPrincipale As Boolean 'indique un affichage complet des outils de la fen�tre principale

'--- pour la configuration ---
Public MemMenuPrincipalNavigateur As Integer   'm�moire de la position du menu principal du navigateur
Public MemSousMenuNavigateur As Integer         'm�moire de la position du sous menu du navigateur

'--- pour l'entretien des fichiers des graphes de production ---
Public EntretienGraphesProduction As Boolean

'--- pour les param�tres du logiciel (cl� "Configuration" dans la base des registres) ---
Public RepLocalBD As String                              'r�pertoire en local des bases de donn�es
Public RepDistantBD As String                           'r�pertoire en distant des bases de donn�es
Public ModeDeConnexion As Integer                  '0=en r�seau, 1=en autonome

Public PARAMETRES_CONNEXION_BD_ANODISATION_SQL As String
Public PARAMETRES_CONNEXION_BD_CLIPPER_HF As String
Public MODE_SECOURS As Boolean

Public SuppressionMotsDePasse As Boolean   'suppression des demandes de mots de passe dans le logiciel
Public MotDePasseDirection As String               'mot de passe de la direction
Public MotDePassePersonnel As String             'mot de passe personnel
Public UniteMonetaire As Integer                        '0=Francs fran�ais, 1=Euros
Public IndicePrestationParDefaut As Integer       'indice de la prestation par d�faut (0=CHROMAGE, etc...)
Public LibellePrestationParDefaut As String       'libell� de la prestation par d�faut ("CHROMAGE", etc...)
Public NbrLignesMaxiAExtraire As Long             'nombre de lignes maxi � extraire pour les grandes tables ou requ�tes

Public DISTANCE_SECURITE As Long                  ' distance de s�curit� pour l'anti-collision
Public DEBUG_MODE  As Boolean

'--- pour les mots de passe ---
Public TypeDeMotDePasse As Boolean
    
'--- entr�e automatique des charges ---
Public EntreeAutomatiqueCharges As Boolean  'entr�e automatique des charges
    
'--- pour le copier / coller sp�cial ---
Public NumFenetreEnCopie As Long                  'indique le num�ro de fenetre en copie pour le copier / coller sp�cial
Public CleDeCopie As Variant                             'indique la cl� de copie (N� de commande interne, N� du devis, ...)

'--- pour l'impression des �tats ---
Public OptionImpressionChoisie As Integer       'option d'impression pour les �tats
Public MargeGaucheTwips As Long                   'marge gauche en twips pour l'impression
Public MargeHauteTwips As Long                      'marge haute en twips pour l'impression
Public MargeDroiteTwips As Long                      'marge droite en twips pour l'impression
Public MargeBasseTwips As Long                     'marge basse en twips pour l'impression
Public PersonneEmettrice As String                   'nom de la personne �mettrice

'--- pour les modifications avant impression ---
Public MemReperefenetreCritereRecherche As String 'm�moire du rep�re de la fenetre d'appel et du crit�re de recherche

'--- caract�res sp�ciaux ---
Public CARACTERE_PHI As String * 1                   'caract�re phi (diam�tre)
Public CARACTERE_FRANC As String * 1              'caract�re pour le franc
Public CARACTERE_EURO As String * 1               'caract�re pour l'euro

'--- pour le filtrage des touches ---
Public ModeSurFrappe As Boolean

'--- limites du tableau des zones ---
Public LIMITE_BASSE_ZONES As Integer             'limite basse du tableau des zones
Public LIMITE_HAUTE_ZONES As Integer             'limite haute du tableau des zones

'--- limites du tableau des barres ---
Public LIMITE_BASSE_BARRES As Integer             'limite basse du tableau des barres
Public LIMITE_HAUTE_BARRES As Integer             'limite haute du tableau des barres
'--- images ---
Public ImgFondDeFenetre As Picture                     'image de fond standard d'un fen�tre
Public ImgFondDeFenetreXP As Picture                 'image de fond standard d'une fen�tre type XP
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

'--- pour le report des d�fauts sur le gyrophare et le klaxon ---
Public SignalerDefautSurGyrophare As Boolean   'TRUE = Indique une demande de d�clenchement du gyrophare
Public SignalerDefautSurKlaxon As Boolean         'TRUE = Indique une demande de d�clenchement du klaxon

'--- pour la gestion de la fin de journ�e ---
Public GestionFinDeJourneeEnCours As Boolean 'TRUE = Indique une demande de fin de journ�e
                                                                                 'FALSE = Gestion normale de la ligne

'--- alarmes de la ligne en cours
'    cette variable contient toutes les alarmes (s�par�es par le s�parateur des alarmes) de la ligne � l'instant x
Public AlarmesLigneEnCours As String

'--- priorit� de l'afficheur pour les alertes ---
Public PrioriteAfficheurPourAlertes As Boolean       'FALSE = pas de priorit� d'affichage des alertes donc affichage des d�fauts
                                                                                  'TRUE = priorit� d'affichage des alertes

'--- temps de compensation d'anodisation en minutes ---
Public TempsCompensationAnodisationMinutes As Integer

'--- mode d'affichage du synoptique et du chargement pr�visionnel ---
'permet l'affichage entre le num�ro de barre et de charge
Public ModeAffichageSynoptique As MODES_AFFICHAGES_SYNOPTIQUE

'*** TABLEAUX ***

'--- images des �l�ments ---
Public TImgEchelles24H(ECHELLES_24H.E_CHAUFFAGE To ECHELLES_24H.E_VENTILATION_CHAUFFAGE) As Picture 'images des �chelles 24 heures pour le programmateur cyclique
Public TRedresseursBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_VERT To IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_EXCLUS) As Picture
Public TRedresseursZoomBMP(IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_VERT To IMAGES_REDRESSEURS.I_BAS_REDRESSEUR_EXCLUS) As Picture

'--- m�moires temporaires des enregistrements des tables pour lecture ou �criture de donn�es ---
Public TTempEnrGammesAnodisation As EnrGammesAnodisation
Public TTempEnrDetailsGammesAnodisation() As EnrDetailsGammesAnodisation
Public TTempEnrCommandesInterne As EnrCommandesInterne
Public TTempEnrDetailsChargesProduction() As EnrDetailsChargesProduction
Public TTempEnrDetailsGammesProduction() As EnrDetailsGammesProduction
Public TTempEnrDetailsPhasesProduction() As EnrDetailsPhasesProduction
Public TTempEnrDetailsFichesProduction() As EnrDetailsFichesProduction
Public TTempEnrPhasesClipper As EnrPhasesClipper

'--- pour l'impression des �tats ---
Public TCriteresImpression(1 To 10) As Variant      'crit�res d'impression (param�tres)

'--- mati�res ---
Public TMatieres(1 To 50) As EnrMatieres

'--- zones ---
Public TZones() As EnrZones                                 'zones de la ligne d'anodisation
'--- barres ---
Public TBarres() As EnrBarres                                 'zones de la ligne d'anodisation

'--- actions ---
Public TActions(NUM_ACTION_NOP To NUM_ACTION_FCY) As EnrActions

'--- image de la m�moire des cycles des ponts (sans adaptation pour les param�tres) ---
Public TImageAPICyclesPonts(PONTS.P_1 To PONTS.P_2, 1 To NBR_LIGNES_CYCLES_PONTS) As Integer

'--- �tats des ponts ---
Public TEtatsPonts(PONTS.P_1 To PONTS.P_2) As EtatsPonts

'--- caract�ristiques des cuves ---
Public TCaracteristiquesCuves(CUVES.C_C00 To CUVES.C_C02) As CaracteristiquesCuves

'--- �tats de la ligne ---
Public TEtatsLigne As EtatsLigne                          '�tats de la ligne

'--- �tats des cuves ---
Public TEtatsCuves(CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As EtatsCuves

'--- �tats des postes ---
Public TEtatsPostes(POSTES.P_CHGT_1 To DERNIER_POSTE) As EtatsPostes

'--- �tats des charges ---
Public TEtatsCharges(CHARGES.C_NUM_MINI To CHARGES.C_NUM_MAXI) As etatsCharges

'--- pr�misses ---
Public TPremisses(POSTES.P_CHGT_1 To DERNIER_POSTE, POSTES.P_CHGT_1 To DERNIER_POSTE) As VarPremisses

'--- chargement ---
Public TChargement As VarChargement

'--- pr�visionnel ---
Public TPrevisionnel(1 To NBR_LIGNES_PREVISIONNEL) As VarPrevisionnel

'--- �tats des redresseurs ---
Public TEtatsRedresseurs(REDRESSEURS.R_C13 To REDRESSEURS.R_C19) As EtatsRedresseurs

'--- �tats des annexes (ventilation, volet de compensation, ...) ---
Public TEtatsAnnexes As EtatsAnnexes

'--- d�fauts ---
Public TDefauts(DEFAUTS.NUM_MINI To DEFAUTS.NUM_MAXI) As EnrDefauts

'--- commandes op�rateur ---
Public TCommandesOperateur(1 To 20) As VarCommandesOperateur  'commandes de l'op�rateur

'--- moteur d'inf�rence ---
Public TMoteurInference As VarMoteurInference      'moteur d'inf�rence (contient toutes les donn�es)


'--- journ�es types ---
Public TJourneesTypes(CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION, JOURNEES_TYPES.J_ARRET To JOURNEES_TYPES.J_REPRISE) As VarCycle24HJourneesTypes

'--- programmateur cyclique ---
Public TProgCyclique(1 To NBR_JOURS_PROG_CYCLIQUE, CUVES_REGULATION.C_C00 To DERNIERE_CUV_REGULATION) As VarCycle24HProgCyclique

'--- m�morisation de divers manipulations ---
Public VManipsGestionRegulation As ManipsGestionRegulation
Public VManipsProgCyclique As ManipsProgCyclique

'--- renseignements sur le graphe � imprimer ---
Public TRenseignementsGraphe As RenseignementsGraphe

'--- tableau des canaux des fichiers de tra�abilit� ---
Public TCanauxFichiersTra�abilite(REDRESSEURS.R_C13 To REDRESSEURS.R_C16) As Integer






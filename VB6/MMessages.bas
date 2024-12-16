Attribute VB_Name = "MMessages"
'---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE CONTENANT LES CONSTANTES ET ROUTINES DES MESSAGES
' Nom                    : MMessages.bas
' Date de création : 07/11/2000
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- pour le déroulement des gammes en automatique ---
Public Const OK  As String = "PAS D'ERREUR"
Public Const PONT_NON_AUTOMATIQUE  As String = "LE PONT N'EST PAS EN AUTOMATIQUE"
Public Const ERREUR_COMMUNICATION_API  As String = "ERREUR DE COMMUNICATION AVEC L'AUTOMATE"
Public Const CYCLE_DEJA_EN_COURS  As String = "CYCLE DEJA EN COURS"
Public Const FIN As String = "FIN"
Public Const PREMISSE_INEXISTANTE As String = "PREMISSE INEXISTANTE"
Public Const MAUVAIS_POSTE_DEPART_ARRIVEE As String = "MAUVAIS POSTE DE DEPART OU D'ARRIVEE"

'--- messages standards ---
Public Const MESSAGE_1 As String = vbCrLf & "Certaines valeurs ont changées." & vbCrLf & vbCrLf & vbCrLf & "cs|Désirez-vous valider ces valeurs avant de quitter ?"
Public Const MESSAGE_2 As String = vbCrLf & vbCrLf & vbCrLf & "Etes-vous sûr de vouloir supprimer cet enregistrement ?"
Public Const MESSAGE_3 As String = vbCrLf & vbCrLf & vbCrLf & "Etes-vous sûr de vouloir supprimer la ligne "
Public Const MESSAGE_4 As String = vbCrLf & vbCrLf & "Etes-vous sûr de vouloir transférer toutes ces valeurs" & vbCrLf & vbCrLf & "dans l'automate ?"
Public Const MESSAGE_5 As String = vbCrLf & "cs|VOTRE RESPONSABILITE EST ENGAGEE" & vbCrLf & vbCrLf & "Etes-vous sûr de vouloir transférer cette valeur de" & vbCrLf & vbCrLf & "POINTEUR dans l'automate ?"
Public Const MESSAGE_6 As String = vbCrLf & vbCrLf & "c|Valeur de POINTEUR NON CONFORME" & vbCrLf & vbCrLf & "c|MINI = 0, MAXI = "
Public Const MESSAGE_7 As String = vbCrLf & vbCrLf & vbCrLf & "cs|SAUVEGARDE effectuée avec succès"
Public Const MESSAGE_8 As String = vbCrLf & vbCrLf & vbCrLf & "cs|CHARGEMENT effectué avec succès"
Public Const MESSAGE_9 As String = vbCrLf & vbCrLf & vbCrLf & "cs|FICHIER NON TROUVE"
Public Const MESSAGE_10 As String = vbCrLf & vbCrLf & vbCrLf & "Transférer la mise en ARRET dans l'automate ?"
Public Const MESSAGE_11 As String = vbCrLf & vbCrLf & vbCrLf & "Transférer la mise en MARCHE dans l'automate ?"
Public Const MESSAGE_12 As String = vbCrLf & vbCrLf & vbCrLf & "Transférer l'EXCLUSION dans l'automate ?"

Public Const MESSAGE_100 As String = "Format d'heure incorrect" & vbCrLf & "Mini 00:00" & vbCrLf & "Maxi 23:59"
Public Const MESSAGE_101 As String = "Format d'heure incorrect" & vbCrLf & "Mini 00:00" & vbCrLf & "Maxi 99:59"
Public Const MESSAGE_105 As String = "Format de date incorrect"
Public Const MESSAGE_106 As String = "Numéro de semaine incorrect"

Public Const MESSAGE_110 As String = "Traitement impossible car division par zéro"

Public Const MESSAGE_118 As String = vbCrLf & vbCrLf & vbCrLf & "c|Ce numéro de commande interne existe déjà" & vbCrLf & "c|dans la liste."
Public Const MESSAGE_119 As String = "Cette valeur ne correspond pas à une personne de l'entreprise"
Public Const MESSAGE_120 As String = "Le numéro d'article tapé n'existe pas"
Public Const MESSAGE_121 As String = vbCrLf & vbCrLf & vbCrLf & "Fiche(s) non trouvée(s)"
Public Const MESSAGE_122 As String = vbCrLf & vbCrLf & vbCrLf & "c|Commande interne non trouvée"
Public Const MESSAGE_123 As String = vbCrLf & vbCrLf & vbCrLf & "Le graphe de production n'a pas été trouvé"
Public Const MESSAGE_124 As String = "Devis non trouvée"
Public Const MESSAGE_125 As String = "Voulez-vous importer cette fiche ?"
Public Const MESSAGE_126 As String = "ATTENTION - Cette fiche est déjà facturée" & vbCrLf & "Confirmez-vous cette saisie ?"
Public Const MESSAGE_127 As String = "ATTENTION - Le prix remis est inférieur au prix total calculé" & vbCrLf & "Confirmez-vous l'enregistrement du devis ?"
Public Const MESSAGE_128 As String = "ATTENTION - Le nombre de pièces est incorrect" & vbCrLf & "Confirmez-vous l'enregistrement de la commande interne ?"
Public Const MESSAGE_129 As String = "Pas de bains trouvés"

Public Const MESSAGE_130 As String = "Numéro de pointage incorrect"
Public Const MESSAGE_131 As String = vbCrLf & vbCrLf & vbCrLf & "c|Gamme inexistante"
Public Const MESSAGE_132 As String = vbCrLf & "Confirmez-vous l'ARRET des éléments suivants :" & vbCrLf & vbCrLf & _
                                                                                "      - L'électro-vanne d'arrivée d'eau de la ligne" & vbCrLf & _
                                                                                "      - Les compresseurs des 2 ponts" & vbCrLf & _
                                                                                "      - Les éclairages des 2 ponts"
Public Const MESSAGE_133 As String = vbCrLf & "Confirmez-vous la MISE EN SERVICE des éléments" & vbCrLf & "suivants :" & vbCrLf & vbCrLf & _
                                                                                "      - L'électro-vanne d'arrivée d'eau de la ligne" & vbCrLf & _
                                                                                "      - Les compresseurs des 2 ponts" & vbCrLf & _
                                                                                "      - Les éclairages des 2 ponts"
Public Const MESSAGE_134 As String = "ATTENTION - Le numéro de la semaine de début" & vbCrLf & "doit être inférieur ou égal au numéro de la semaine de fin"

Public Const MESSAGE_140 As String = vbCrLf & vbCrLf & "Pas de charge en cours" & vbCrLf & vbCrLf & "Sélectionner une charge afin d'accèder aux options"

Public Const MESSAGE_200 As String = "L'aperçu d'impression n'est pas disponible pour ce choix" & vbCrLf & "Lancer l'impression directement"
Public Const MESSAGE_210 As String = "Valider vos données avant de lancer l'impression"
Public Const MESSAGE_220 As String = "Mauvaise sélection du bon de livraison ou erreur de données"

Public Const MESSAGE_300 As String = vbCrLf & vbCrLf & vbCrLf & "c|Etes-vous sûr de vouloir supprimer le" & vbCrLf & "c|CHARGEMENT ?"
Public Const MESSAGE_301 As String = vbCrLf & vbCrLf & vbCrLf & "c|Etes-vous sûr de vouloir supprimer le" & vbCrLf & "c|PREVISIONNEL ?"
Public Const MESSAGE_302 As String = vbCrLf & "cs|ATTENTION" & vbCrLf & vbCrLf & "c|Etes-vous sur de vouloir supprimer cette charge ?"

Public Const MESSAGE_400 As String = vbCrLf & "Les couleurs ne sont pas valides." & vbCrLf & vbCrLf & vbCrLf & "La résolution graphique doit être configurée en 24 bits" & vbCrLf & "ou plus"

'--- messages relatifs à l'automate ---
Public Const MESSAGE_500 As String = vbCrLf & vbCrLf & "cs|ERREUR SUR LE RESEAU AUTOMATE" & vbCrLf & vbCrLf & vbCrLf & "cs|Les valeurs n'ont pas été tranférées"
Public Const MESSAGE_501 As String = vbCrLf & vbCrLf & "Vous devez choisir un poste."
'--- messages relatifs à l'ordinateur ---
Public Const MESSAGE_550 As String = vbCrLf & vbCrLf & "cs|S'agit-il d'une reprise après un incident" & vbCrLf & vbCrLf & "cs|INFORMATIQUE ou d'AUTOMATISME ?"

'--- messages relatifs au prémisses et IA ---
Public Const MESSAGE_600 As String = vbCrLf & "cs|ATTENTION" & vbCrLf & vbCrLf & "c|Etes-vous sûr de vouloir supprimer cette prémisse ?"
Public Const MESSAGE_601 As String = vbCrLf & "cs|ATTENTION" & vbCrLf & vbCrLf & "Etes-vous sûr de vouloir une REGENERATION COMPLETE" & vbCrLf & "des PREMISSES ?"

'--- messages relatifs à la gestion des redresseurs ---
Public Const MESSAGE_700 As String = vbCrLf & "cs|ATTENTION" & vbCrLf & vbCrLf & "c|Etes-vous sûr de vouloir arrêter ce redresseur ?"
Public Const MESSAGE_701 As String = vbCrLf & "cs|ATTENTION" & vbCrLf & vbCrLf & "c|Etes-vous sûr de vouloir passer le redresseur" & vbCrLf & "c|en attente de la dépose d'une charge ?"


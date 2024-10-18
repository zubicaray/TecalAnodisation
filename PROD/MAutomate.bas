Attribute VB_Name = "MAutomate"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle                    : MODULE GERANT LE DIALOGUE AVEC L'AUTOMATE
' Nom                    : MAutomate.bas
' Date de création : 14/11/2003
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Détail de la fonction READ : Public Function Read (ByVal Source As Long,
'                                                                                   ByVal NbItems As Long,
'                                                                                   ByVal ItemsRef As Variant,
'                                                                                   ByRef Value As Variant,
'                                                                                   ByRef Quality As Variant,
'                                                                                   ByRef TimeStamp As Variant,
'                                                                                   ByRef TabStatus As Variant ) As Long
'
' Valeur retournée par la fonction :
'                                Elle retourne l'état de la communication entre l'activeX et le serveur OPC.
'                                Valeur retournée -> 0 = La communication s'est bien passée
'
'                                                               < 0 Problème de communication (voir ci-dessous)
'
'                                                               -1 = Nom du serveur invalide
'                                                               -2 = Nom du groupe invalide
'                                                               -3 = Nom de l'item invalide
'                                                               -4 = Référence sur un serveur invalide (dans toutes les fonctions sauf GetGroupRef)
'                                                               -5 = Référence sur un groupe invalide
'                                                               -6 = Référence sur un item invalide
'                                                               -7 = Paramètre invalide
'                                                               -8 = Liste vide
'                                                               -9 = Erreur sur la référence d'un serveur lors d'une demande de référence sur un groupe
'                                                             -10 = Erreur sur la référence d'un groupe lors d'une demande de référence sur un item
'                                                             -11 = Erreur sur la référence d'un item
'                                                             -12 = Référence invalide
'                                                             -13 = Erreur sur l'écriture
'                                                             -14 = Erreur sur la lecture
'                                                             -15 = La variable itemRef n'est pas un entier 32 bits
'                                                             -17 = Le nombre d'items est nul (égal à 0)
'                                                             -21 = L'écriture est impossible : problème de communication avec le serveur OPC
'                                                             -22 = Le chargement de la configuration est impossible
'                                                             -24 = L'item n'a pas pu être ajouté dans le serveur OPC
'                                                                      Cette erreur peut être retournée par la fonction ActiveConfig,
'                                                                      si vous utilisez l'applicom® communication ActiveX control avec un serveur OPC applicom®
'                                                                      dans le cadre d'une solution SW1000ETH et que votre configuration dépasse le nombre d'items
'                                                                      autorisé par la protection logicielle
'                                                             -25 = Une erreur s'est produite lors de l'accès au fichier configopc.mdb
'                                                             -26 = Le groupe n'a pas pu être ajouté dans le serveur OPC
'                                                             -27 = Le fichier appActivex.log n'a pas pu être créé
'                                                             -28 = Problème de connexion avec le serveur
'                                                             -29 = Le serveur est déjà dans l'état demandé
'                                                             -30 = Le serveur OPC applicom est absent ou non actif
'                                                             -31 = Nom de l'item invalide
'                                                             -32 = Référence sur un serveur non présent dans la configuration
'                                                             -33 = Référence sur un groupe non présent dans la configuration
'                                                             -34 = Référence sur un item non présent dans la configuration
'                                                             -35 = Connexion impossible à au moins un serveur OPC de la configuration
'                                                             -36 = La base de configuration n'est pas au bon format
'                                                           -249 = Pas de configuration activée
'
' Paramètres en entrée Type :
'                           Source -> Entier 32 bits, indiquant la source des données
'                                           0 pour lire dans la mémoire cache, 1 pour lire dans l'équipement
'                         NbItems -> Entier 32 bits, indiquant le nombre d'items (membres) à lire
'                       I temsRef -> Variant de type Entier 32 bits (VT_I4), ou tableau d'entiers 32 bits (VT_ARRAY|VT_I4)
'                                            contenant les références sur un ou des items. Ces paramètres sont retournés par la fonction GetItemRef ()
'
' Paramètres en sortie Type :
'                             Value -> Variant de type tableau de Variant (VT_ARRAY|VT_VARIANT) contenant la valeur de chaque item (membre)
'                                           Ce Variant dépend du type de données de chaque item.
'                           Quality -> Variant de type tableau d'entiers 32 bits (VT_ARRAY|VT_I4) contenant la qualité de chaque item (membre)
'                                           Les valeurs possibles sont :    0 = Qualité mauvaise
'                                                                                           28 = Qualité inactive
'                                                                                           64 = Qualité incertaine
'                                                                                         192 = Qualité bonne
'                    TimeStamp -> Variant de type tableau d'entiers 32 bits (VT_ARRAY|VT_I4) contenant l'horodatage de la valeur
'                                           pour chaque item (membre)
'                      TabStatus -> Variant de type tableau d'entiers 32 bits (VT_ARRAY|VT_I4) contenant l'état de la communication
'                                           pour chaque item (membre)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- déclarations obligatoires ---
Option Explicit

'--- options générales ---
Option Base 1
DefVar A-Z

'--- constantes publiques concernat la librairie de communication Applicom ---
Public Const NOM_SERVEUR_APPLICOM = "ANODISATION"           'nom du serveur Applicom
Public Const APP_256BYTES_BASED_LIMITS = 256
Public Const APP_1584BYTES_BASED_LIMITS = 1584

'--- variable publiques ---
Public RefServeur As Long                        'référence sur le serveur applicom

'--- fonctions d'initialisation ---
Declare Sub initbus Lib "applicom" (stat As Integer)
Declare Sub exitbus Lib "applicom" (stat As Integer)
Declare Function AuSetApplicationMaxSize Lib "applicom" (ByVal buffersize As Integer, stat As Long) As Boolean
Declare Function AuGetApplicomMaxSize Lib "applicom" (sizeDB As Integer, sizeChan As Integer, stat As Long) As Boolean

'--- fonctions en mode cyclique ---
Declare Sub createcyc Lib "applicom" (chan As Integer, fonc As Integer, per As Integer, act As Integer, typf As Integer, codedb As Integer, nes As Integer, nb As Integer, adrl As Long, adrDB As Integer, adrstat As Integer, stat As Integer)
Declare Sub actcyc Lib "applicom" (chan As Integer, fonc As Integer, stat As Integer)
Declare Sub startcyc Lib "applicom" (chan As Integer, fonc As Integer, stat As Integer)
Declare Sub stopcyc Lib "applicom" (chan As Integer, fonc As Integer, stat As Integer)
Declare Sub transcyc Lib "applicom" (chan As Integer, fonc As Integer, nb As Integer, typ As Integer, tabl As Long, stat As Integer)
Declare Sub transcycpack Lib "applicom" (chan As Integer, fonc As Integer, nb As Integer, typ As Integer, tabl As Integer, stat As Integer)
Declare Sub CycExecuted Lib "applicom" (chan As Integer, nb As Integer, talb As Integer, stat As Integer)
Declare Sub AppGetCycParam Lib "applicom" (chan As Integer, fonc As Integer, tabl As Long, stat As Integer)

'--- fonctions de la base de données ---
Declare Sub confdb Lib "applicom" (card As Integer, typ As Integer, nb As Integer, tadr As Integer, cond As Integer, stat As Integer)
Declare Sub getevent Lib "applicom" (card As Integer, nb As Integer, nestyp As Integer, nesAdr As Integer, nesval As Long, nesdat As Integer, nestime As Integer, stat As Integer)
Declare Sub setevent Lib "applicom" (card As Integer, typ As Integer, nb As Integer, adr As Integer, stat As Integer)
Declare Sub incdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub incdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub decdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub decdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub getdispbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub getdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub getdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub getdispfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Single, stat As Integer)
Declare Sub setdispbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub setdispword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Integer, stat As Integer)
Declare Sub setdispdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Long, stat As Integer)
Declare Sub setdispfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tval As Single, stat As Integer)
Declare Sub getbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub getpackbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub getpackbyte Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Byte, stat As Integer)
Declare Sub getword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub getdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Long, stat As Integer)
Declare Sub getfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Single, stat As Integer)
Declare Sub setbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub setpackbit Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub setpackbyte Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Byte, stat As Integer)
Declare Sub setword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub setdword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Long, stat As Integer)
Declare Sub setfword Lib "applicom" (card As Integer, nb As Integer, tadr As Integer, tabl As Single, stat As Integer)
Declare Sub invbit Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)

'--- fonctions en mode attente ---
Declare Sub readquickbit Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readpackbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readpackibit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readpackqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readpackbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub readpackibyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub readpackqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub readbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readibyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readwordbcd Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readiword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub readdword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Long, stat As Integer)
Declare Sub readfword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Single, stat As Integer)

Declare Sub writepackbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writepackqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writepackbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Byte, stat As Integer)
Declare Sub writepackqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub writebyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writeqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writeword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writewordbcd Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writeqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Long, stat As Integer)
Declare Sub writefword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Single, stat As Integer)
Declare Sub readcounter Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub writecounter Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub readtimer Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub writetimer Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, adr As Long, tabl As Integer, stat As Integer)
Declare Sub readmes Lib "applicom" (chan As Integer, tstop As Integer, nb As Integer, tim As Integer, tim_ic As Integer, tabl As Byte, stat As Integer)
Declare Sub writemes Lib "applicom" (chan As Integer, nb As Integer, tabl As Byte, stat As Integer)
Declare Sub writereadmes Lib "applicom" (chan As Integer, nb_tx As Integer, buf_tx As Byte, tstop As Integer, max_rx As Integer, time_out As Integer, time_ic As Integer, nb_rx As Integer, buf_rx As Byte, stat As Integer)

'--- fonctions en mode différé ---
Declare Sub readdifquickbit Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub readdifbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Integer, stat As Integer)
Declare Sub readdifibit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Integer, stat As Integer)
Declare Sub readdifibyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifiword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdifdword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub readdiffword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, stat As Integer)
Declare Sub writedifpackbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifpackqbit Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifpackbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub writedifpackqbyte Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Byte, stat As Integer)
Declare Sub writedifword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifqword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Integer, stat As Integer)
Declare Sub writedifdword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Long, stat As Integer)
Declare Sub writediffword Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tadr As Long, tabl As Single, stat As Integer)
Declare Sub readdifmes Lib "applicom" (chan As Integer, tstop As Integer, nb As Integer, tim As Integer, tim_ic As Integer, stat As Integer)
Declare Sub writedifmes Lib "applicom" (chan As Integer, nb As Integer, tabl As Byte, stat As Integer)
Declare Sub writereaddifmes Lib "applicom" (chan As Integer, nbtx As Integer, buftx As Byte, tstop As Integer, nbrx As Integer, timeout As Integer, timeic As Integer, stat As Integer)
Declare Sub transdif Lib "applicom" (chan As Integer, nb As Integer, typ As Integer, tabl As Long, stat As Integer)
Declare Sub transdifpack Lib "applicom" (chan As Integer, nb As Integer, typ As Integer, tabl As Integer, stat As Integer)
Declare Sub testtransdif Lib "applicom" (chan As Integer, request As Integer, receive As Integer, stat As Integer)

'--- fonctions spécifiques UTE ---
Declare Sub readcounterute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub writecounterute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, preset As Integer, stat As Integer)
Declare Sub resetcounterute Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub readiobitute Lib "applicom" (chan As Integer, nes As Integer, module As Integer, tabl As Integer, status As Integer)
Declare Sub writeiobitute Lib "applicom" (chan As Integer, nes As Integer, modu As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub readmonostableute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub writemonostableute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, preset As Integer, stat As Integer)
Declare Sub readtimerute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, tabl As Integer, stat As Integer)
Declare Sub writetimerute Lib "applicom" (chan As Integer, nes As Integer, tadr As Integer, preset As Integer, stat As Integer)
Declare Sub readdiagute Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub txtute Lib "applicom" (chan As Integer, nes As Integer, txtic As Integer, txil As Integer, bufemi As Byte, bufrec As Byte, txtis As Integer, txtiv As Integer, stat As Integer)

'--- fonctions diverses ---
Declare Sub AppConnect Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub AppUnconnect Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub bcdbin Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub binbcd Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub getmodem Lib "applicom" (chan As Integer, cts As Integer, dcd As Integer, stat As Integer)
Declare Sub setmodem Lib "applicom" (chan As Integer, rts As Integer, dtr As Integer, stat As Integer)
Declare Sub transbitword Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub transwordbit Lib "applicom" (nb As Integer, tabls As Integer, tabld As Integer, stat As Integer)
Declare Sub unpackdate Lib "applicom" (pdate As Integer, dday As Integer, mmonth As Integer, yyear As Integer, stat As Integer)
Declare Sub unpacktime Lib "applicom" (ptime As Integer, min As Integer, seg As Integer, hhour As Integer, stat As Integer)
'Declare Sub watchdog Lib "applicom" (card As Integer, tim As Integer, stat As Integer)
'Declare Sub AppGetWatchDog Lib "applicom" (card As Integer, Entree As Integer, Contact As Integer, stat As Integer)
Declare Sub iocounter Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub resetiocounter Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub readdiag Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readeven Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readtrace Lib "applicom" (chan As Integer, nes As Integer, tabl As Integer, stat As Integer)
Declare Sub readident Lib "applicom" (chan As Integer, nes As Integer, nb As Integer, tabl As Byte, stat As Integer)
Declare Sub accesskey Lib "applicom" (chan As Integer, rts As Integer, dtr As Integer, stat As Integer)
Declare Sub automatic Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub compword Lib "applicom" (nb As Integer, tabls1 As Integer, tabls2 As Integer, tabld As Integer, stat As Integer)
Declare Sub manual Lib "applicom" (chan As Integer, nes As Integer, stat As Integer)
Declare Sub statjbus Lib "applicom" (chan As Integer, tabl As Integer, stat As Integer)
Declare Sub AppFmsGetOd Lib "applicom" (chan As Integer, nes As Integer, attribut As Integer, Acces As Integer, Index As Integer, nb As Integer, tabl As Byte, More As Integer, stat As Integer)
Declare Sub AppFmsStatus Lib "applicom" (chan As Integer, nes As Integer, logStat As Integer, PhyStat As Integer, stat As Integer)
Declare Sub AppIniTime Lib "applicom" (card As Integer, stat As Integer)

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture de la chaine de caractères représentant les messages sur la qualité des membres
' Entrées : Qualite -> valeur représentant la qualité d'un membre lors d'une lecture ou écriture
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function MessagesQualite(ByVal Qualite As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case Qualite

        Case 0
            '--- mauvais ---
            MessagesQualite = "Mauvais"
    
        Case 28
            '--- groupe inactif ---
            MessagesQualite = "Groupe inactif"
    
        Case 64
            '--- incertain ---
            MessagesQualite = "Incertain"
    
        Case 192
            '--- bon ---
            MessagesQualite = "Bon"
    
        Case Else
    End Select

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Lecture de la chaine de caractères représentant les messages sur l'état de la communication
' Entrées : EtatCommunication -> valeur représentant l'état de la communication
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function MessagesEtatCommunication(ByVal EtatCommunication As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case EtatCommunication
        Case 0: MessagesEtatCommunication = OK 'la communication s'est bien passée
        Case -1: MessagesEtatCommunication = "Nom du serveur invalide"
        Case -2: MessagesEtatCommunication = "Nom du groupe invalide"
        Case -3: MessagesEtatCommunication = "Nom de l'item invalide"
        Case -4: MessagesEtatCommunication = "Référence sur un serveur invalide (dans toutes les fonctions sauf GetGroupRef)"
        Case -5: MessagesEtatCommunication = "Référence sur un groupe invalide"
        Case -6: MessagesEtatCommunication = "Référence sur un item invalide"
        Case -7: MessagesEtatCommunication = "Paramètre invalide"
        Case -8: MessagesEtatCommunication = "Liste vide"
        Case -9: MessagesEtatCommunication = "Erreur sur la référence d'un serveur lors d'une demande de référence sur un groupe"
        Case -10: MessagesEtatCommunication = "Erreur sur la référence d'un groupe lors d'une demande de référence sur un item"
        Case -11: MessagesEtatCommunication = "Erreur sur la référence d'un item"
        Case -12: MessagesEtatCommunication = "Référence invalide"
        Case -13: MessagesEtatCommunication = "Erreur sur l'écriture"
        Case -14: MessagesEtatCommunication = "Erreur sur la lecture"
        Case -15: MessagesEtatCommunication = "La variable itemRef n'est pas un entier 32 bits"
        Case -17: MessagesEtatCommunication = "Le nombre d'items est nul (égal à 0)"
        Case -21: MessagesEtatCommunication = "L'écriture est impossible : problème de communication avec le serveur OPC"
        Case -22: MessagesEtatCommunication = "Le chargement de la configuration est impossible"
        Case -24: MessagesEtatCommunication = "L'item n'a pas pu être ajouté dans le serveur OPC"
                                                                            'cette erreur peut être retournée par la fonction ActiveConfig,"
                                                                            'si vous utilisez l'applicom® communication ActiveX control avec un serveur OPC applicom®
                                                                            'dans le cadre d'une solution SW1000ETH et que votre configuration dépasse le nombre d'items
                                                                            'autorisé par la protection logicielle
        Case -25: MessagesEtatCommunication = "Une erreur s'est produite lors de l'accès au fichier configopc.mdb"
        Case -26: MessagesEtatCommunication = "Le groupe n'a pas pu être ajouté dans le serveur OPC"
        Case -27: MessagesEtatCommunication = "Le fichier appActivex.log n'a pas pu être créé"
        Case -28: MessagesEtatCommunication = "Problème de connexion avec le serveur"
        Case -29: MessagesEtatCommunication = "Le serveur est déjà dans l'état demandé"
        Case -30: MessagesEtatCommunication = "Le serveur OPC applicom est absent ou non actif"
        Case -31: MessagesEtatCommunication = "Nom de l'item invalide"
        Case -32: MessagesEtatCommunication = "Référence sur un serveur non présent dans la configuration"
        Case -33: MessagesEtatCommunication = "Référence sur un groupe non présent dans la configuration"
        Case -34: MessagesEtatCommunication = "Référence sur un item non présent dans la configuration"
        Case -35: MessagesEtatCommunication = "Connexion impossible à au moins un serveur OPC de la configuration"
        Case -36: MessagesEtatCommunication = "La base de configuration n'est pas au bon format"
        Case -249: MessagesEtatCommunication = "Pas de configuration activée"
        Case Else: MessagesEtatCommunication = ERREUR_COMMUNICATION_API
    End Select

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Extraction de la chaine de caractères représentant l'heure d'une valeur d'horodatage
' Entrées : Horodatage -> Valeur d'horodatage retournée par une fonction de communication
'                                        (lecture ou écriture)
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function HeureHorodatage(ByVal Horodatage As Date) As String
    On Error Resume Next
    HeureHorodatage = Mid(Horodatage, 12, 8)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Extraction de la chaine de caractères représentant la date d'une valeur d'horodatage
' Entrées : Horodatage -> Valeur d'horodatage retournée par une fonction de communication
'                                        (lecture ou écriture)
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DateHorodatage(ByVal Horodatage As Date) As String
    On Error Resume Next
    DateHorodatage = Mid(Horodatage, 1, 10)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Active une configuration
' Entrées : AOCConcerne -> Objet applicom AppOcxClient concerné
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ActiveConfiguration(ByRef AOCConcerne As AppOcxClient) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- déclaration ---
    
    '--- activation de la configuration ---
    ActiveConfiguration = AOCConcerne.ActiveConfig

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture sur une variable nommée à partir de l'outil OCX APPLICOM
' Entrées :                          NomGroupe   -> Nom du groupe (voir notice APPLICOM)
'                                          NomVariable -> Nom de la variable (voir notice APPLICOM)
'                                                    Valeur -> Valeur à transmettre
' Retours : APIEcritureVariableNommee -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureVariableNommee(ByVal NomGroupe As String, _
                                                                        ByVal NomVariable As String, _
                                                                        ByVal Valeur As Variant) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    'Call Log("APIEcritureVariableNommee", "NomGroupe:" & NomGroupe & " ;NomVariable=" & NomVariable & " ; valeur=" & CStr(Valeur))
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim RefGroupe As Long                         'référence sur le groupe
    Dim RefElement As Long                        'référence sur un élément
    Dim TEtatsCommunication As Variant    'tableau contenant les états de la communication de chaque valeur transmise
    
    '--- référence sur le groupe ---
    RefGroupe = OccFPrincipale.AOCFPrincipale.GetGroupRef(RefServeur, NomGroupe)
    If RefGroupe <= 0 Then
        APIEcritureVariableNommee = RefGroupe
        Exit Function
    End If
            
    '--- écriture du numéro d'OF ---
    RefElement = OccFPrincipale.AOCFPrincipale.GetItemRef(RefGroupe, NomVariable)
    If RefElement <= 0 Then
        APIEcritureVariableNommee = RefElement
        Exit Function
    End If
    
    '--- écriture ---
    APIEcritureVariableNommee = OccFPrincipale.AOCFPrincipale.Write(1, RefElement, Valeur, TEtatsCommunication)

    Exit Function

GestionErreurs:
    APIEcritureVariableNommee = -28

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture sur une variable nommée à partir de l'outil OCX APPLICOM
' Entrées :                              NomGroupe -> Nom du groupe (voir notice APPLICOM)
'                                            EtatSouhaite ->   TRUE = activation du groupe
'                                                                      FALSE = désactivation du groupe
' Retours : APIActiveDesactiveUnGroupe -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIActiveDesactiveUnGroupe(ByVal NomGroupe As String, _
                                                                          ByVal EtatSouhaite As Boolean) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes privées ---
    
    '--- déclaration ---
    Dim RefGroupe As Long                         'référence sur le groupe
    Dim TEtatsCommunication As Variant    'tableau contenant les états de la communication de chaque valeur transmise
    
    '--- référence sur le groupe ---
    RefGroupe = OccFPrincipale.AOCFPrincipale.GetGroupRef(RefServeur, NomGroupe)
    If RefGroupe <= 0 Then
        APIActiveDesactiveUnGroupe = RefGroupe
        Exit Function
    End If
        
    '--- envoi de la commande ---
    APIActiveDesactiveUnGroupe = OccFPrincipale.AOCFPrincipale.SetGroupState(RefGroupe, EtatSouhaite)
    
    Exit Function

GestionErreurs:
    APIActiveDesactiveUnGroupe = -28

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture sur le bit
' Entrées :         LibelleBit -> Libellé du bit (exemple : M123.2)
' Retours :         ValeurBit -> Valeur du bit à écrire
'                 APIEcritureBit -> 0 = Transmission bonne, sinon numéro de l'erreur
' Retours :
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureBit(ByVal LibelleBit As String, _
                                                 ByVal ValeurBit As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APIEcritureBit", "LibelleBit:" & LibelleBit & " ;ValeurBit=" & ValeurBit)
    '--- constantes privées ---
        
    '--- déclaration ---
    Dim NbrBits As Integer                         'nombre de bits
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer      'état de la communication
    Dim AdresseBit As Long                       'adresse du bit
    Dim TValeurBit(1) As Integer                 'tableau contenant la valeur du bit
    Dim TDetailsAdresse As Variant           'tableau contenanrt les détails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 à 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrBits = 1
    TValeurBit(1) = ValeurBit
    
    '--- extraction des détails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleBit, Len(LibelleBit) - 1), ".")
    
    '--- détermination de l'adresse ---
    AdresseBit = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- écriture ---
    Call writepackbit(NumCanal, NumEquipement, NbrBits, AdresseBit, TValeurBit(1), EtatCommunication)

    '--- valeur de retour ---
    APIEcritureBit = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Monte le bit de défaut pour appeler l'opérateur
' Entrées :
' Retours : APIAppelOperateur -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIAppelOperateur() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const M_DEM_PC_KLAXON As String = "M12.0"
    Const M_DEM_PC_VOYANTS As String = "M12.4"
    
    '--- déclaration ---

    '--- transfert des valeurs ---
    If PROGRAMME_AVEC_AUTOMATE = True Then

        '--- écriture du bit dans l'automate ---
        APIAppelOperateur = APIEcritureBit(M_DEM_PC_KLAXON, 1)
        If APIAppelOperateur <> 0 Then Exit Function
        
        APIAppelOperateur = APIEcritureBit(M_DEM_PC_VOYANTS, 1)

    End If

End Function
                                                
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la lecture d'un bit dans l'automate
' Entrées :        LibelleBit -> Libellé du bit (exemple : M123.2)
' Retours :        ValeurBit -> Valeur du bit après lecture
'                 APILectureBit -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureBit(ByVal LibelleBit As String, _
                                                ByRef ValeurBit As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    ''Call Log("APILectureBit", "LibelleBit:" & LibelleBit & " ;ValeurBit=" & CStr(ValeurBit))
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrBits As Integer                         'nombre de bits
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer      'état de la communication
    Dim AdresseBit As Long                       'adresse du bit
    Dim TValeurBit(1) As Integer                 'tableau contenant la valeur du bit
    Dim TDetailsAdresse As Variant           'tableau contenanrt les détails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 à 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrBits = 1

    '--- extraction des détails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleBit, Len(LibelleBit) - 1), ".")
    
    '--- détermination de l'adresse ---
    AdresseBit = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- lecture du bit ---
    Call readpackbit(NumCanal, NumEquipement, NbrBits, AdresseBit, TValeurBit(1), EtatCommunication)
        
    '--- valeur de retour ---
    ValeurBit = TValeurBit(1)
    APILectureBit = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la lecture d'une entrée de l'automate
' Entrées :        LibelleEntree -> Libellé de l'entrée (exemple : E10.2)
' Retours :        ValeurEntree -> Valeur de l'entrée après lecture
'                 APILectureEntree -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureEntree(ByVal LibelleEntree As String, _
                                                       ByRef ValeurEntree As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APILectureEntree", "LibelleEntree:" & LibelleEntree & " ;ValeurEntree=" & CStr(ValeurEntree))
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrEntrees As Integer                    'nombre d'entrées
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer      'état de la communication
    Dim AdresseEntree As Long                 'adresse de l'entrée
    Dim TValeurEntree(1) As Integer           'tableau contenant la valeur de l'entrée
    Dim TDetailsAdresse As Variant           'tableau contenanrt les détails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 à 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrEntrees = 1

    '--- extraction des détails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleEntree, Len(LibelleEntree) - 1), ".")
    
    '--- détermination de l'adresse ---
    AdresseEntree = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- lecture ---
    Call readpackibit(NumCanal, NumEquipement, NbrEntrees, AdresseEntree, TValeurEntree(1), EtatCommunication)
        
    '--- valeur de retour ---
    ValeurEntree = TValeurEntree(1)
    APILectureEntree = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la lecture d'une sortie de l'automate
' Entrées :         LibelleSortie -> Libellé de la sortie (exemple : A31.2)
' Retours :         ValeurSortie -> Valeur de la sortie après lecture
'                 APILectureEntree -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureSortie(ByVal LibelleSortie As String, _
                                                      ByRef ValeurSortie As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APILectureSortie", "LibelleSortie:" & LibelleSortie & " ;ValeurSortie=" & CStr(ValeurSortie))
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrSorties As Integer                    'nombre de sorties
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer      'état de la communication
    Dim AdresseSortie As Long                  'adresse de la sortie
    Dim TValeurSortie(1) As Integer           'tableau contenant la valeur de la sortie
    Dim TDetailsAdresse As Variant           'tableau contenanrt les détails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 à 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrSorties = 1

    '--- extraction des détails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleSortie, Len(LibelleSortie) - 1), ".")
    
    '--- détermination de l'adresse ---
    AdresseSortie = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- lecture ---
    Call readpackqbit(NumCanal, NumEquipement, NbrSorties, AdresseSortie, TValeurSortie(1), EtatCommunication)
        
    '--- valeur de retour ---
    ValeurSortie = TValeurSortie(1)
    APILectureSortie = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la lecture d'un octet dans l'automate
' Entrées :        LibelleOctet -> Libellé du bit (exemple : MB10)
' Retours :        ValeurOctet -> Valeur de l'octet après lecture
'                 APILectureOctet -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureOctet(ByVal LibelleOctet As String, _
                                                     ByRef ValeurOctet As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APILectureOctet", "LibelleOctet:" & LibelleOctet & " ;ValeurOctet=" & CStr(ValeurOctet))
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrOctets As Integer                      'nombre d'octets
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer      'état de la communication
    Dim AdresseOctet As Long                   'adresse de l'octet
    Dim TValeurOctet(1) As Byte                'tableau contenant la valeur de l'octet
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrOctets = 1
    
    '--- détermination de l'adresse ---
    AdresseOctet = CLng(Right(LibelleOctet, Len(LibelleOctet) - 2))

    '--- lecture ---
    Call readpackbyte(NumCanal, NumEquipement, NbrOctets, AdresseOctet, TValeurOctet(1), EtatCommunication)

    '--- valeur de retour ---
    ValeurOctet = TValeurOctet(1)
    APILectureOctet = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la lecture d'un double mots dans l'automate
' Entrées :        LibelleDoubleMots -> Libellé du double mots (exemple : MD123)
' Retours :        ValeurDoubleMots -> Valeur du double mots après lecture
'                 APILectureDoubleMots -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureDoubleMots(ByVal LibelleDoubleMots As String, _
                                                               ByRef ValeurDoubleMots As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APILectureDoubleMots", "LibelleDoubleMots:" & LibelleDoubleMots & " ;ValeurDoubleMots=" & CStr(ValeurDoubleMots))
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrDoublesMots As Integer           'nombre de doubles mots
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer       'état de la communication
    Dim AdresseDoubleMots As Long         'adresse du double mots
    Dim TValeurDoubleMots(1) As Long      'tableau contenant le double mots
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrDoublesMots = 1
    
    '--- détermination de l'adresse ---
    AdresseDoubleMots = CLng(Right(LibelleDoubleMots, Len(LibelleDoubleMots) - 2))

    '--- lecture ---
    Call readdword(NumCanal, NumEquipement, NbrDoublesMots, AdresseDoubleMots, TValeurDoubleMots(1), EtatCommunication)
    
    '--- valeur de retour ---
    ValeurDoubleMots = TValeurDoubleMots(1)
    APILectureDoubleMots = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la lecture d'un flottant (double mot de type décimale) dans l'automate
' Entrées :        LibelleFlottant -> Libellé du flottant (exemple : MD123)
' Retours :        ValeurFlottant -> Valeur du flottant après lecture
'                 APILectureFlottant -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureFlottant(ByVal LibelleFlottant As String, _
                                                        ByRef ValeurFlottant As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APILectureFlottant", "LibelleFlottant:" & LibelleFlottant & " ;ValeurFlottant=" & CStr(ValeurFlottant))
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrFlottants As Integer                  'nombre de flottant
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer       'état de la communication
    Dim AdresseFlottant As Long                'adresse du flottant
    Dim TValeurFlottant(1) As Single           'tableau contenant le flottant
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrFlottants = 1
    
    '--- détermination de l'adresse ---
    AdresseFlottant = CLng(Right(LibelleFlottant, Len(LibelleFlottant) - 2))

    '--- lecture ---
    Call readfword(NumCanal, NumEquipement, NbrFlottants, AdresseFlottant, TValeurFlottant(1), EtatCommunication)
    
    '--- valeur de retour ---
    ValeurFlottant = TValeurFlottant(1)
    APILectureFlottant = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une lecture sur le mot
' Entrées :        LibelleMot -> Libellé du Mot (exemple : MW100)
' Retours :          ValeurBit -> Valeur du mot après lecture
'                 APILectureMot -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureMot(ByVal LibelleMot As String, _
                                                  ByRef ValeurMot As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APILectureMot", "LibelleMot:" & LibelleMot & " ;ValeurMot=" & CStr(ValeurMot))
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrMots As Integer                        'nombre de mots
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer       'état de la communication
    Dim AdresseMot As Long                      'adresse du mot
    Dim TValeurMot(1) As Integer                'tableau contenant la valeur du mot
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrMots = 1
    
    '--- détermination de l'adresse ---
    AdresseMot = CLng(Right(LibelleMot, Len(LibelleMot) - 2))

    '--- lecture ---
    Call readword(NumCanal, NumEquipement, NbrMots, AdresseMot, TValeurMot(1), EtatCommunication)
    
    '--- valeur de retour ---
    ValeurMot = TValeurMot(1)
    APILectureMot = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture d'un mot dans l'automate
' Entrées :        LibelleMot -> Libellé du Mot (exemple : MW100)
'                        ValeurMot -> Valeur du mot à écrire
' Retours : APILectureMot -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureMot(ByVal LibelleMot As String, _
                                                   ByVal ValeurMot As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes privées ---
    'Call Log("APIEcritureMot", "LibelleMot:" & LibelleMot & " ;ValeurMot=" & ValeurMot)
    '--- déclaration ---
    Dim NbrMots As Integer                               'nombre de mots
    Dim NumCanal As Integer                            'numéro du canal
    Dim NumEquipement As Integer                  'numéro de l'équipement
    Dim EtatCommunication As Integer             'état de la communication
    Dim AdresseMot As Long                            'adresse du mot
    Dim TValeurMot(1) As Integer                      'tableau contenant la valeur du mot
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrMots = 1
    TValeurMot(1) = ValeurMot
    
    '--- détermination de l'adresse ---
    AdresseMot = CLng(Right(LibelleMot, Len(LibelleMot) - 2))
    
    '--- écriture ---
    Call writeword(NumCanal, NumEquipement, NbrMots, AdresseMot, TValeurMot(1), EtatCommunication)

    '--- valeur de retour ---
    APIEcritureMot = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture d'un double mots dans l'automate
' Entrées :        LibelleDoubleMots -> Libellé du double mots (exemple : MD100)
'                        ValeurDoubleMots -> Valeur du double mots à écrire
' Retours : APIEcritureDoubleMots -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureDoubleMots(ByVal LibelleDoubleMots As String, _
                                                                ByVal ValeurDoubleMots As Long) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APIEcritureDoubleMots", "LibelleDoubleMots:" & LibelleDoubleMots & " ;ValeurDoubleMots=" & ValeurDoubleMots)
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrDoublesMots As Integer           'nombre de doubles mots
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer       'état de la communication
    Dim AdresseDoubleMots As Long         'adresse du double mots
    Dim TValeurDoubleMots(1) As Long      'tableau contenant le double mots
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrDoublesMots = 1
    
    '--- détermination de l'adresse ---
    AdresseDoubleMots = CLng(Right(LibelleDoubleMots, Len(LibelleDoubleMots) - 2))

    '--- écriture ---
    Call writedword(NumCanal, NumEquipement, NbrDoublesMots, AdresseDoubleMots, TValeurDoubleMots(1), EtatCommunication)
    
    '--- valeur de retour ---
    APIEcritureDoubleMots = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue l'écriture d'une charge dans l'automate
' Entrées :             NumCharge -> Numéro de charge
'                          TEtatsCharge -> Tableau des états d'une charge
' Retours : APITransfertCharge -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APITransfertCharge(ByVal NumCharge As Integer, _
                                                          ByRef TEtatsCharge As etatsCharges) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    'PtrZoneGammeAnodisation
    Dim msg As String
    Dim aa As Integer
    
    msg = ""
    
   ' For aa = PREMIER_BAIN To DERNIER_POSTE


    '    msg = msg & "TEtatsPostes(" & aa & ").PtrZoneGammeAnodisation = " & TEtatsCharges(aa).PtrZoneGammeAnodisation & Chr(13)
    '    msg = msg & "TEtatsPostes(" & aa & ").NumBarre = " & TEtatsCharges(aa).NumBarre & Chr(13)
        
      
    'Next aa
    
    'Call Log(msg)
    
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim a As Integer                                                        'pour les boucles FOR...NEXT
    Dim NbrDonneesTotalATransmettre As Integer        'nombre de données total à transmettre à l'automate
    Dim NbrDonneesTransmisesAPI As Integer             'nombre de données transmises à l'automate
    
    Dim Cle As String                                                      'représente une clé pour une recherche unique
    Dim NomGroupe As String                                        'représente un nom de groupe

    '--- affectation du nom du groupe ---
    NomGroupe = "CHARGE_" & Right("00" & NumCharge, 2)
    
    '--- affectation par défaut ---
    APITransfertCharge = 0
    
    '--- affectation des valeurs de comptage des données transmises à l'automate ---
    NbrDonneesTotalATransmettre = 16            'nombre de données total à transmettre à l'automate
    NbrDonneesTransmisesAPI = 0                   'nombre de données transmises à l'automate
    
    With TEtatsCharge
        
        '************************************************************************************************************************************************
        '                                                                             TRANSFERT DE LA FICHE DE BASE
        '************************************************************************************************************************************************
    
        '--- numéro de commande interne ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre

        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "NumCommandeInterne", Replace(.TDetailsCharges(1).NumCommandeInterne, "C", ""))
        If APITransfertCharge <> 0 Then Exit Function
    
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- numéro de la barre ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "NumBarre", .NumBarre)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- mode U ou I ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "ModeUouI", .ModeUouI)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- temps de la phase 1 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "TpsPhase1", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1).TempsPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- tension de la phase 1 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "UPhase1", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1).UPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- intensité de la phase 1 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "IPhase1", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T1).IPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- temps de la phase 2 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "TpsPhase2", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T2).TempsPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- tension de la phase 2 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "UPhase2", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T2).UPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- intensité de la phase 2 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "IPhase2", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T2).IPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- temps de la phase 3 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "TpsPhase3", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T3).TempsPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- tension de la phase 3 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "UPhase3", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T3).UPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- intensité de la phase 3 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "IPhase3", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T3).IPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- temps de la phase 4 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "TpsPhase4", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).TempsPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- tension de la phase 4 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "UPhase4", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).UPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- intensité de la phase 4 ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "IPhase4", .TDetailsPhasesProduction(PHASES_GAMMES_REDRESSEURS.PH_T4).IPhase)
        If APITransfertCharge <> 0 Then Exit Function
        
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- temps total de la gamme ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre
        
        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "TempsTotalCycle", .TempsTotalGammeRedresseur)
        If APITransfertCharge <> 0 Then Exit Function
        
    End With

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture d'un flottant (double mot de type décimale) dans l'automate
' Entrées :        LibelleFlottant -> Libellé du flottant (exemple : MD123)
'                        ValeurFlottant -> Valeur du flottant à écrire
' Retours : APIEcritureFlottant -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureFlottant(ByVal LibelleFlottant As String, _
                                                        ByVal ValeurFlottant As Single) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APIEcritureFlottant", "LibelleFlottant:" & LibelleFlottant & " ;ValeurFlottant=" & ValeurFlottant)
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrFlottants As Integer                  'nombre de flottants
    Dim NumCanal As Integer                     'numéro du canal
    Dim NumEquipement As Integer           'numéro de l'équipement
    Dim EtatCommunication As Integer       'état de la communication
    Dim AdresseFlottant As Long                'adresse du flottant
    Dim TValeurFlottant(1) As Single           'tableau contenant le flottant
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrFlottants = 1
    TValeurFlottant(1) = ValeurFlottant
    
    '--- détermination de l'adresse ---
    AdresseFlottant = CLng(Right(LibelleFlottant, Len(LibelleFlottant) - 2))

    '--- écriture ---
    Call writefword(NumCanal, NumEquipement, NbrFlottants, AdresseFlottant, TValeurFlottant(1), EtatCommunication)
    
    '--- valeur de retour ---
    APIEcritureFlottant = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue une écriture d'un octet dans l'automate
' Entrées :        LibelleOctet -> Libellé du bit (exemple : MB10)
'                        ValeurOctet -> Valeur de l'octet à écrire
' Retours : APILectureOctet -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureOctet(ByVal LibelleOctet As String, _
                                                     ByVal ValeurOctet As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APIEcritureOctet", "LibelleOctet:" & LibelleOctet & " ;ValeurOctet=" & ValeurOctet)
    '--- constantes privées ---
    
    '--- déclaration ---
    Dim NbrOctets As Integer                             'nombre d'octets
    Dim NumCanal As Integer                            'numéro du canal
    Dim NumEquipement As Integer                  'numéro de l'équipement
    Dim EtatCommunication As Integer             'état de la communication
    Dim AdresseOctet As Long                          'adresse de l'octet
    Dim TValeurOctet(1) As Byte                        'tableau contenant la valeur de l'octet
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrOctets = 1
    TValeurOctet(1) = ValeurOctet
    
    '--- détermination de l'adresse ---
    AdresseOctet = CLng(Right(LibelleOctet, Len(LibelleOctet) - 2))
    
    '--- écriture ---
    Call writepackbyte(NumCanal, NumEquipement, NbrOctets, AdresseOctet, TValeurOctet(1), EtatCommunication)

    '--- valeur de retour ---
    APIEcritureOctet = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Effectue la condamnation ou décondamnation d'un poste
' Entrées :                                                 NumPoste -> Numéro d'un poste
'                                                             EtatSouhaite -> FALSE = Décondamnation
'                                                                                        TRUE = Condamnation
' Retours : APICondamnationDecondamnationPoste -> 0 = Transmission bonne, sinon numéro de l'erreur
' Détails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APICondamnationDecondamnationPoste(ByVal NumPoste As POSTES, _
                                                                                           ByVal EtatSouhaite As Boolean) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes privées ---
    Const NOM_GROUPE = "SUIVI_LIGNE"    'nom du groupe
    
    '--- déclaration ---
    Dim NomVariable As String                       'nom de la variable OPC
        
    '--- affectation du nom de la variable ---
    NomVariable = "CondamnationPoste" & Right("00" & CStr(NumPoste), 2)
                    
    '--- écriture dans l'automate ---
    APICondamnationDecondamnationPoste = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, IIf(EtatSouhaite = False, "0", "1"))

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' Rôle      : Affichage des données transmises à l'API
' Entrées :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AffichageDonneesTransmisesAPI(ByVal TexteAAfficher As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- transmission dans la Fenetre principale
    OccFPrincipale.LDonneesTransmisesAPI = TexteAAfficher

End Sub


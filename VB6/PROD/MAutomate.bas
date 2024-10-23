Attribute VB_Name = "MAutomate"
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le                    : MODULE GERANT LE DIALOGUE AVEC L'AUTOMATE
' Nom                    : MAutomate.bas
' Date de cr�ation : 14/11/2003
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' D�tail de la fonction READ : Public Function Read (ByVal Source As Long,
'                                                                                   ByVal NbItems As Long,
'                                                                                   ByVal ItemsRef As Variant,
'                                                                                   ByRef Value As Variant,
'                                                                                   ByRef Quality As Variant,
'                                                                                   ByRef TimeStamp As Variant,
'                                                                                   ByRef TabStatus As Variant ) As Long
'
' Valeur retourn�e par la fonction :
'                                Elle retourne l'�tat de la communication entre l'activeX et le serveur OPC.
'                                Valeur retourn�e -> 0 = La communication s'est bien pass�e
'
'                                                               < 0 Probl�me de communication (voir ci-dessous)
'
'                                                               -1 = Nom du serveur invalide
'                                                               -2 = Nom du groupe invalide
'                                                               -3 = Nom de l'item invalide
'                                                               -4 = R�f�rence sur un serveur invalide (dans toutes les fonctions sauf GetGroupRef)
'                                                               -5 = R�f�rence sur un groupe invalide
'                                                               -6 = R�f�rence sur un item invalide
'                                                               -7 = Param�tre invalide
'                                                               -8 = Liste vide
'                                                               -9 = Erreur sur la r�f�rence d'un serveur lors d'une demande de r�f�rence sur un groupe
'                                                             -10 = Erreur sur la r�f�rence d'un groupe lors d'une demande de r�f�rence sur un item
'                                                             -11 = Erreur sur la r�f�rence d'un item
'                                                             -12 = R�f�rence invalide
'                                                             -13 = Erreur sur l'�criture
'                                                             -14 = Erreur sur la lecture
'                                                             -15 = La variable itemRef n'est pas un entier 32 bits
'                                                             -17 = Le nombre d'items est nul (�gal � 0)
'                                                             -21 = L'�criture est impossible : probl�me de communication avec le serveur OPC
'                                                             -22 = Le chargement de la configuration est impossible
'                                                             -24 = L'item n'a pas pu �tre ajout� dans le serveur OPC
'                                                                      Cette erreur peut �tre retourn�e par la fonction ActiveConfig,
'                                                                      si vous utilisez l'applicom� communication ActiveX control avec un serveur OPC applicom�
'                                                                      dans le cadre d'une solution SW1000ETH et que votre configuration d�passe le nombre d'items
'                                                                      autoris� par la protection logicielle
'                                                             -25 = Une erreur s'est produite lors de l'acc�s au fichier configopc.mdb
'                                                             -26 = Le groupe n'a pas pu �tre ajout� dans le serveur OPC
'                                                             -27 = Le fichier appActivex.log n'a pas pu �tre cr��
'                                                             -28 = Probl�me de connexion avec le serveur
'                                                             -29 = Le serveur est d�j� dans l'�tat demand�
'                                                             -30 = Le serveur OPC applicom est absent ou non actif
'                                                             -31 = Nom de l'item invalide
'                                                             -32 = R�f�rence sur un serveur non pr�sent dans la configuration
'                                                             -33 = R�f�rence sur un groupe non pr�sent dans la configuration
'                                                             -34 = R�f�rence sur un item non pr�sent dans la configuration
'                                                             -35 = Connexion impossible � au moins un serveur OPC de la configuration
'                                                             -36 = La base de configuration n'est pas au bon format
'                                                           -249 = Pas de configuration activ�e
'
' Param�tres en entr�e Type :
'                           Source -> Entier 32 bits, indiquant la source des donn�es
'                                           0 pour lire dans la m�moire cache, 1 pour lire dans l'�quipement
'                         NbItems -> Entier 32 bits, indiquant le nombre d'items (membres) � lire
'                       I temsRef -> Variant de type Entier 32 bits (VT_I4), ou tableau d'entiers 32 bits (VT_ARRAY|VT_I4)
'                                            contenant les r�f�rences sur un ou des items. Ces param�tres sont retourn�s par la fonction GetItemRef ()
'
' Param�tres en sortie Type :
'                             Value -> Variant de type tableau de Variant (VT_ARRAY|VT_VARIANT) contenant la valeur de chaque item (membre)
'                                           Ce Variant d�pend du type de donn�es de chaque item.
'                           Quality -> Variant de type tableau d'entiers 32 bits (VT_ARRAY|VT_I4) contenant la qualit� de chaque item (membre)
'                                           Les valeurs possibles sont :    0 = Qualit� mauvaise
'                                                                                           28 = Qualit� inactive
'                                                                                           64 = Qualit� incertaine
'                                                                                         192 = Qualit� bonne
'                    TimeStamp -> Variant de type tableau d'entiers 32 bits (VT_ARRAY|VT_I4) contenant l'horodatage de la valeur
'                                           pour chaque item (membre)
'                      TabStatus -> Variant de type tableau d'entiers 32 bits (VT_ARRAY|VT_I4) contenant l'�tat de la communication
'                                           pour chaque item (membre)
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'--- d�clarations obligatoires ---
Option Explicit

'--- options g�n�rales ---
Option Base 1
DefVar A-Z

'--- constantes publiques concernat la librairie de communication Applicom ---
Public Const NOM_SERVEUR_APPLICOM = "ANODISATION"           'nom du serveur Applicom
Public Const APP_256BYTES_BASED_LIMITS = 256
Public Const APP_1584BYTES_BASED_LIMITS = 1584

'--- variable publiques ---
Public RefServeur As Long                        'r�f�rence sur le serveur applicom

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

'--- fonctions de la base de donn�es ---
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

'--- fonctions en mode diff�r� ---
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

'--- fonctions sp�cifiques UTE ---
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
' R�le      : Lecture de la chaine de caract�res repr�sentant les messages sur la qualit� des membres
' Entr�es : Qualite -> valeur repr�sentant la qualit� d'un membre lors d'une lecture ou �criture
' Retours :
' D�tails  :
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
' R�le      : Lecture de la chaine de caract�res repr�sentant les messages sur l'�tat de la communication
' Entr�es : EtatCommunication -> valeur repr�sentant l'�tat de la communication
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Function MessagesEtatCommunication(ByVal EtatCommunication As Long) As String
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    Select Case EtatCommunication
        Case 0: MessagesEtatCommunication = OK 'la communication s'est bien pass�e
        Case -1: MessagesEtatCommunication = "Nom du serveur invalide"
        Case -2: MessagesEtatCommunication = "Nom du groupe invalide"
        Case -3: MessagesEtatCommunication = "Nom de l'item invalide"
        Case -4: MessagesEtatCommunication = "R�f�rence sur un serveur invalide (dans toutes les fonctions sauf GetGroupRef)"
        Case -5: MessagesEtatCommunication = "R�f�rence sur un groupe invalide"
        Case -6: MessagesEtatCommunication = "R�f�rence sur un item invalide"
        Case -7: MessagesEtatCommunication = "Param�tre invalide"
        Case -8: MessagesEtatCommunication = "Liste vide"
        Case -9: MessagesEtatCommunication = "Erreur sur la r�f�rence d'un serveur lors d'une demande de r�f�rence sur un groupe"
        Case -10: MessagesEtatCommunication = "Erreur sur la r�f�rence d'un groupe lors d'une demande de r�f�rence sur un item"
        Case -11: MessagesEtatCommunication = "Erreur sur la r�f�rence d'un item"
        Case -12: MessagesEtatCommunication = "R�f�rence invalide"
        Case -13: MessagesEtatCommunication = "Erreur sur l'�criture"
        Case -14: MessagesEtatCommunication = "Erreur sur la lecture"
        Case -15: MessagesEtatCommunication = "La variable itemRef n'est pas un entier 32 bits"
        Case -17: MessagesEtatCommunication = "Le nombre d'items est nul (�gal � 0)"
        Case -21: MessagesEtatCommunication = "L'�criture est impossible : probl�me de communication avec le serveur OPC"
        Case -22: MessagesEtatCommunication = "Le chargement de la configuration est impossible"
        Case -24: MessagesEtatCommunication = "L'item n'a pas pu �tre ajout� dans le serveur OPC"
                                                                            'cette erreur peut �tre retourn�e par la fonction ActiveConfig,"
                                                                            'si vous utilisez l'applicom� communication ActiveX control avec un serveur OPC applicom�
                                                                            'dans le cadre d'une solution SW1000ETH et que votre configuration d�passe le nombre d'items
                                                                            'autoris� par la protection logicielle
        Case -25: MessagesEtatCommunication = "Une erreur s'est produite lors de l'acc�s au fichier configopc.mdb"
        Case -26: MessagesEtatCommunication = "Le groupe n'a pas pu �tre ajout� dans le serveur OPC"
        Case -27: MessagesEtatCommunication = "Le fichier appActivex.log n'a pas pu �tre cr��"
        Case -28: MessagesEtatCommunication = "Probl�me de connexion avec le serveur"
        Case -29: MessagesEtatCommunication = "Le serveur est d�j� dans l'�tat demand�"
        Case -30: MessagesEtatCommunication = "Le serveur OPC applicom est absent ou non actif"
        Case -31: MessagesEtatCommunication = "Nom de l'item invalide"
        Case -32: MessagesEtatCommunication = "R�f�rence sur un serveur non pr�sent dans la configuration"
        Case -33: MessagesEtatCommunication = "R�f�rence sur un groupe non pr�sent dans la configuration"
        Case -34: MessagesEtatCommunication = "R�f�rence sur un item non pr�sent dans la configuration"
        Case -35: MessagesEtatCommunication = "Connexion impossible � au moins un serveur OPC de la configuration"
        Case -36: MessagesEtatCommunication = "La base de configuration n'est pas au bon format"
        Case -249: MessagesEtatCommunication = "Pas de configuration activ�e"
        Case Else: MessagesEtatCommunication = ERREUR_COMMUNICATION_API
    End Select

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extraction de la chaine de caract�res repr�sentant l'heure d'une valeur d'horodatage
' Entr�es : Horodatage -> Valeur d'horodatage retourn�e par une fonction de communication
'                                        (lecture ou �criture)
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function HeureHorodatage(ByVal Horodatage As Date) As String
    On Error Resume Next
    HeureHorodatage = Mid(Horodatage, 12, 8)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Extraction de la chaine de caract�res repr�sentant la date d'une valeur d'horodatage
' Entr�es : Horodatage -> Valeur d'horodatage retourn�e par une fonction de communication
'                                        (lecture ou �criture)
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function DateHorodatage(ByVal Horodatage As Date) As String
    On Error Resume Next
    DateHorodatage = Mid(Horodatage, 1, 10)
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Active une configuration
' Entr�es : AOCConcerne -> Objet applicom AppOcxClient concern�
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function ActiveConfiguration(ByRef AOCConcerne As AppOcxClient) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- d�claration ---
    
    '--- activation de la configuration ---
    ActiveConfiguration = AOCConcerne.ActiveConfig

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue une �criture sur une variable nomm�e � partir de l'outil OCX APPLICOM
' Entr�es :                          NomGroupe   -> Nom du groupe (voir notice APPLICOM)
'                                          NomVariable -> Nom de la variable (voir notice APPLICOM)
'                                                    Valeur -> Valeur � transmettre
' Retours : APIEcritureVariableNommee -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureVariableNommee(ByVal NomGroupe As String, _
                                                                        ByVal NomVariable As String, _
                                                                        ByVal Valeur As Variant) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    'Call Log("APIEcritureVariableNommee", "NomGroupe:" & NomGroupe & " ;NomVariable=" & NomVariable & " ; valeur=" & CStr(Valeur))
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim RefGroupe As Long                         'r�f�rence sur le groupe
    Dim RefElement As Long                        'r�f�rence sur un �l�ment
    Dim TEtatsCommunication As Variant    'tableau contenant les �tats de la communication de chaque valeur transmise
    
    '--- r�f�rence sur le groupe ---
    RefGroupe = OccFPrincipale.AOCFPrincipale.GetGroupRef(RefServeur, NomGroupe)
    If RefGroupe <= 0 Then
        APIEcritureVariableNommee = RefGroupe
        Exit Function
    End If
            
    '--- �criture du num�ro d'OF ---
    RefElement = OccFPrincipale.AOCFPrincipale.GetItemRef(RefGroupe, NomVariable)
    If RefElement <= 0 Then
        APIEcritureVariableNommee = RefElement
        Exit Function
    End If
    
    '--- �criture ---
    APIEcritureVariableNommee = OccFPrincipale.AOCFPrincipale.Write(1, RefElement, Valeur, TEtatsCommunication)

    Exit Function

GestionErreurs:
    APIEcritureVariableNommee = -28

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue une �criture sur une variable nomm�e � partir de l'outil OCX APPLICOM
' Entr�es :                              NomGroupe -> Nom du groupe (voir notice APPLICOM)
'                                            EtatSouhaite ->   TRUE = activation du groupe
'                                                                      FALSE = d�sactivation du groupe
' Retours : APIActiveDesactiveUnGroupe -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIActiveDesactiveUnGroupe(ByVal NomGroupe As String, _
                                                                          ByVal EtatSouhaite As Boolean) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error GoTo GestionErreurs

    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim RefGroupe As Long                         'r�f�rence sur le groupe
    Dim TEtatsCommunication As Variant    'tableau contenant les �tats de la communication de chaque valeur transmise
    
    '--- r�f�rence sur le groupe ---
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
' R�le      : Effectue une �criture sur le bit
' Entr�es :         LibelleBit -> Libell� du bit (exemple : M123.2)
' Retours :         ValeurBit -> Valeur du bit � �crire
'                 APIEcritureBit -> 0 = Transmission bonne, sinon num�ro de l'erreur
' Retours :
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureBit(ByVal LibelleBit As String, _
                                                 ByVal ValeurBit As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APIEcritureBit", "LibelleBit:" & LibelleBit & " ;ValeurBit=" & ValeurBit)
    '--- constantes priv�es ---
        
    '--- d�claration ---
    Dim NbrBits As Integer                         'nombre de bits
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer      '�tat de la communication
    Dim AdresseBit As Long                       'adresse du bit
    Dim TValeurBit(1) As Integer                 'tableau contenant la valeur du bit
    Dim TDetailsAdresse As Variant           'tableau contenanrt les d�tails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 � 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrBits = 1
    TValeurBit(1) = ValeurBit
    
    '--- extraction des d�tails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleBit, Len(LibelleBit) - 1), ".")
    
    '--- d�termination de l'adresse ---
    AdresseBit = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- �criture ---
    Call writepackbit(NumCanal, NumEquipement, NbrBits, AdresseBit, TValeurBit(1), EtatCommunication)

    '--- valeur de retour ---
    APIEcritureBit = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Monte le bit de d�faut pour appeler l'op�rateur
' Entr�es :
' Retours : APIAppelOperateur -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIAppelOperateur() As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const M_DEM_PC_KLAXON As String = "M12.0"
    Const M_DEM_PC_VOYANTS As String = "M12.4"
    
    '--- d�claration ---

    '--- transfert des valeurs ---
    If PROGRAMME_AVEC_AUTOMATE = True Then

        '--- �criture du bit dans l'automate ---
        APIAppelOperateur = APIEcritureBit(M_DEM_PC_KLAXON, 1)
        If APIAppelOperateur <> 0 Then Exit Function
        
        APIAppelOperateur = APIEcritureBit(M_DEM_PC_VOYANTS, 1)

    End If

End Function
                                                
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la lecture d'un bit dans l'automate
' Entr�es :        LibelleBit -> Libell� du bit (exemple : M123.2)
' Retours :        ValeurBit -> Valeur du bit apr�s lecture
'                 APILectureBit -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureBit(ByVal LibelleBit As String, _
                                                ByRef ValeurBit As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    ''Call Log("APILectureBit", "LibelleBit:" & LibelleBit & " ;ValeurBit=" & CStr(ValeurBit))
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrBits As Integer                         'nombre de bits
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer      '�tat de la communication
    Dim AdresseBit As Long                       'adresse du bit
    Dim TValeurBit(1) As Integer                 'tableau contenant la valeur du bit
    Dim TDetailsAdresse As Variant           'tableau contenanrt les d�tails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 � 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrBits = 1

    '--- extraction des d�tails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleBit, Len(LibelleBit) - 1), ".")
    
    '--- d�termination de l'adresse ---
    AdresseBit = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- lecture du bit ---
    Call readpackbit(NumCanal, NumEquipement, NbrBits, AdresseBit, TValeurBit(1), EtatCommunication)
        
    '--- valeur de retour ---
    ValeurBit = TValeurBit(1)
    APILectureBit = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la lecture d'une entr�e de l'automate
' Entr�es :        LibelleEntree -> Libell� de l'entr�e (exemple : E10.2)
' Retours :        ValeurEntree -> Valeur de l'entr�e apr�s lecture
'                 APILectureEntree -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureEntree(ByVal LibelleEntree As String, _
                                                       ByRef ValeurEntree As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APILectureEntree", "LibelleEntree:" & LibelleEntree & " ;ValeurEntree=" & CStr(ValeurEntree))
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrEntrees As Integer                    'nombre d'entr�es
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer      '�tat de la communication
    Dim AdresseEntree As Long                 'adresse de l'entr�e
    Dim TValeurEntree(1) As Integer           'tableau contenant la valeur de l'entr�e
    Dim TDetailsAdresse As Variant           'tableau contenanrt les d�tails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 � 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrEntrees = 1

    '--- extraction des d�tails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleEntree, Len(LibelleEntree) - 1), ".")
    
    '--- d�termination de l'adresse ---
    AdresseEntree = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- lecture ---
    Call readpackibit(NumCanal, NumEquipement, NbrEntrees, AdresseEntree, TValeurEntree(1), EtatCommunication)
        
    '--- valeur de retour ---
    ValeurEntree = TValeurEntree(1)
    APILectureEntree = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la lecture d'une sortie de l'automate
' Entr�es :         LibelleSortie -> Libell� de la sortie (exemple : A31.2)
' Retours :         ValeurSortie -> Valeur de la sortie apr�s lecture
'                 APILectureEntree -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureSortie(ByVal LibelleSortie As String, _
                                                      ByRef ValeurSortie As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APILectureSortie", "LibelleSortie:" & LibelleSortie & " ;ValeurSortie=" & CStr(ValeurSortie))
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrSorties As Integer                    'nombre de sorties
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer      '�tat de la communication
    Dim AdresseSortie As Long                  'adresse de la sortie
    Dim TValeurSortie(1) As Integer           'tableau contenant la valeur de la sortie
    Dim TDetailsAdresse As Variant           'tableau contenanrt les d�tails de l'adresse
                                                                  'index 0 = position de l'octet
                                                                  'index 1 = position du bit dans l'octet (0 � 7)
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrSorties = 1

    '--- extraction des d�tails de l'adresse ---
    TDetailsAdresse = Split(Right(LibelleSortie, Len(LibelleSortie) - 1), ".")
    
    '--- d�termination de l'adresse ---
    AdresseSortie = TDetailsAdresse(0) * 8 + TDetailsAdresse(1)
    
    '--- lecture ---
    Call readpackqbit(NumCanal, NumEquipement, NbrSorties, AdresseSortie, TValeurSortie(1), EtatCommunication)
        
    '--- valeur de retour ---
    ValeurSortie = TValeurSortie(1)
    APILectureSortie = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la lecture d'un octet dans l'automate
' Entr�es :        LibelleOctet -> Libell� du bit (exemple : MB10)
' Retours :        ValeurOctet -> Valeur de l'octet apr�s lecture
'                 APILectureOctet -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureOctet(ByVal LibelleOctet As String, _
                                                     ByRef ValeurOctet As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    ''Call Log("APILectureOctet", "LibelleOctet:" & LibelleOctet & " ;ValeurOctet=" & CStr(ValeurOctet))
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrOctets As Integer                      'nombre d'octets
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer      '�tat de la communication
    Dim AdresseOctet As Long                   'adresse de l'octet
    Dim TValeurOctet(1) As Byte                'tableau contenant la valeur de l'octet
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrOctets = 1
    
    '--- d�termination de l'adresse ---
    AdresseOctet = CLng(Right(LibelleOctet, Len(LibelleOctet) - 2))

    '--- lecture ---
    Call readpackbyte(NumCanal, NumEquipement, NbrOctets, AdresseOctet, TValeurOctet(1), EtatCommunication)

    '--- valeur de retour ---
    ValeurOctet = TValeurOctet(1)
    APILectureOctet = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la lecture d'un double mots dans l'automate
' Entr�es :        LibelleDoubleMots -> Libell� du double mots (exemple : MD123)
' Retours :        ValeurDoubleMots -> Valeur du double mots apr�s lecture
'                 APILectureDoubleMots -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureDoubleMots(ByVal LibelleDoubleMots As String, _
                                                               ByRef ValeurDoubleMots As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APILectureDoubleMots", "LibelleDoubleMots:" & LibelleDoubleMots & " ;ValeurDoubleMots=" & CStr(ValeurDoubleMots))
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrDoublesMots As Integer           'nombre de doubles mots
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer       '�tat de la communication
    Dim AdresseDoubleMots As Long         'adresse du double mots
    Dim TValeurDoubleMots(1) As Long      'tableau contenant le double mots
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrDoublesMots = 1
    
    '--- d�termination de l'adresse ---
    AdresseDoubleMots = CLng(Right(LibelleDoubleMots, Len(LibelleDoubleMots) - 2))

    '--- lecture ---
    Call readdword(NumCanal, NumEquipement, NbrDoublesMots, AdresseDoubleMots, TValeurDoubleMots(1), EtatCommunication)
    
    '--- valeur de retour ---
    ValeurDoubleMots = TValeurDoubleMots(1)
    APILectureDoubleMots = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la lecture d'un flottant (double mot de type d�cimale) dans l'automate
' Entr�es :        LibelleFlottant -> Libell� du flottant (exemple : MD123)
' Retours :        ValeurFlottant -> Valeur du flottant apr�s lecture
'                 APILectureFlottant -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureFlottant(ByVal LibelleFlottant As String, _
                                                        ByRef ValeurFlottant As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APILectureFlottant", "LibelleFlottant:" & LibelleFlottant & " ;ValeurFlottant=" & CStr(ValeurFlottant))
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrFlottants As Integer                  'nombre de flottant
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer       '�tat de la communication
    Dim AdresseFlottant As Long                'adresse du flottant
    Dim TValeurFlottant(1) As Single           'tableau contenant le flottant
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrFlottants = 1
    
    '--- d�termination de l'adresse ---
    AdresseFlottant = CLng(Right(LibelleFlottant, Len(LibelleFlottant) - 2))

    '--- lecture ---
    Call readfword(NumCanal, NumEquipement, NbrFlottants, AdresseFlottant, TValeurFlottant(1), EtatCommunication)
    
    '--- valeur de retour ---
    ValeurFlottant = TValeurFlottant(1)
    APILectureFlottant = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue une lecture sur le mot
' Entr�es :        LibelleMot -> Libell� du Mot (exemple : MW100)
' Retours :          ValeurBit -> Valeur du mot apr�s lecture
'                 APILectureMot -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APILectureMot(ByVal LibelleMot As String, _
                                                  ByRef ValeurMot As Variant) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APILectureMot", "LibelleMot:" & LibelleMot & " ;ValeurMot=" & CStr(ValeurMot))
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrMots As Integer                        'nombre de mots
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer       '�tat de la communication
    Dim AdresseMot As Long                      'adresse du mot
    Dim TValeurMot(1) As Integer                'tableau contenant la valeur du mot
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrMots = 1
    
    '--- d�termination de l'adresse ---
    AdresseMot = CLng(Right(LibelleMot, Len(LibelleMot) - 2))

    '--- lecture ---
    Call readword(NumCanal, NumEquipement, NbrMots, AdresseMot, TValeurMot(1), EtatCommunication)
    
    '--- valeur de retour ---
    ValeurMot = TValeurMot(1)
    APILectureMot = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue une �criture d'un mot dans l'automate
' Entr�es :        LibelleMot -> Libell� du Mot (exemple : MW100)
'                        ValeurMot -> Valeur du mot � �crire
' Retours : APILectureMot -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureMot(ByVal LibelleMot As String, _
                                                   ByVal ValeurMot As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    
    '--- constantes priv�es ---
    'Call Log("APIEcritureMot", "LibelleMot:" & LibelleMot & " ;ValeurMot=" & ValeurMot)
    '--- d�claration ---
    Dim NbrMots As Integer                               'nombre de mots
    Dim NumCanal As Integer                            'num�ro du canal
    Dim NumEquipement As Integer                  'num�ro de l'�quipement
    Dim EtatCommunication As Integer             '�tat de la communication
    Dim AdresseMot As Long                            'adresse du mot
    Dim TValeurMot(1) As Integer                      'tableau contenant la valeur du mot
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrMots = 1
    TValeurMot(1) = ValeurMot
    
    '--- d�termination de l'adresse ---
    AdresseMot = CLng(Right(LibelleMot, Len(LibelleMot) - 2))
    
    '--- �criture ---
    Call writeword(NumCanal, NumEquipement, NbrMots, AdresseMot, TValeurMot(1), EtatCommunication)

    '--- valeur de retour ---
    APIEcritureMot = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue une �criture d'un double mots dans l'automate
' Entr�es :        LibelleDoubleMots -> Libell� du double mots (exemple : MD100)
'                        ValeurDoubleMots -> Valeur du double mots � �crire
' Retours : APIEcritureDoubleMots -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureDoubleMots(ByVal LibelleDoubleMots As String, _
                                                                ByVal ValeurDoubleMots As Long) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APIEcritureDoubleMots", "LibelleDoubleMots:" & LibelleDoubleMots & " ;ValeurDoubleMots=" & ValeurDoubleMots)
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrDoublesMots As Integer           'nombre de doubles mots
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer       '�tat de la communication
    Dim AdresseDoubleMots As Long         'adresse du double mots
    Dim TValeurDoubleMots(1) As Long      'tableau contenant le double mots
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrDoublesMots = 1
    
    '--- d�termination de l'adresse ---
    AdresseDoubleMots = CLng(Right(LibelleDoubleMots, Len(LibelleDoubleMots) - 2))

    '--- �criture ---
    Call writedword(NumCanal, NumEquipement, NbrDoublesMots, AdresseDoubleMots, TValeurDoubleMots(1), EtatCommunication)
    
    '--- valeur de retour ---
    APIEcritureDoubleMots = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue l'�criture d'une charge dans l'automate
' Entr�es :             NumCharge -> Num�ro de charge
'                          TEtatsCharge -> Tableau des �tats d'une charge
' Retours : APITransfertCharge -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
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
    
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim a As Integer                                                        'pour les boucles FOR...NEXT
    Dim NbrDonneesTotalATransmettre As Integer        'nombre de donn�es total � transmettre � l'automate
    Dim NbrDonneesTransmisesAPI As Integer             'nombre de donn�es transmises � l'automate
    
    Dim Cle As String                                                      'repr�sente une cl� pour une recherche unique
    Dim NomGroupe As String                                        'repr�sente un nom de groupe

    '--- affectation du nom du groupe ---
    NomGroupe = "CHARGE_" & Right("00" & NumCharge, 2)
    
    '--- affectation par d�faut ---
    APITransfertCharge = 0
    
    '--- affectation des valeurs de comptage des donn�es transmises � l'automate ---
    NbrDonneesTotalATransmettre = 16            'nombre de donn�es total � transmettre � l'automate
    NbrDonneesTransmisesAPI = 0                   'nombre de donn�es transmises � l'automate
    
    With TEtatsCharge
        
        '************************************************************************************************************************************************
        '                                                                             TRANSFERT DE LA FICHE DE BASE
        '************************************************************************************************************************************************
    
        '--- num�ro de commande interne ---
        Inc NbrDonneesTransmisesAPI
        AffichageDonneesTransmisesAPI NbrDonneesTransmisesAPI & " / " & NbrDonneesTotalATransmettre

        APITransfertCharge = APIEcritureVariableNommee(NomGroupe, "NumCommandeInterne", Replace(.TDetailsCharges(1).NumCommandeInterne, "C", ""))
        If APITransfertCharge <> 0 Then Exit Function
    
        '------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        '--- num�ro de la barre ---
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
        '--- intensit� de la phase 1 ---
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
        '--- intensit� de la phase 2 ---
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
        '--- intensit� de la phase 3 ---
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
        '--- intensit� de la phase 4 ---
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
' R�le      : Effectue une �criture d'un flottant (double mot de type d�cimale) dans l'automate
' Entr�es :        LibelleFlottant -> Libell� du flottant (exemple : MD123)
'                        ValeurFlottant -> Valeur du flottant � �crire
' Retours : APIEcritureFlottant -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureFlottant(ByVal LibelleFlottant As String, _
                                                        ByVal ValeurFlottant As Single) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APIEcritureFlottant", "LibelleFlottant:" & LibelleFlottant & " ;ValeurFlottant=" & ValeurFlottant)
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrFlottants As Integer                  'nombre de flottants
    Dim NumCanal As Integer                     'num�ro du canal
    Dim NumEquipement As Integer           'num�ro de l'�quipement
    Dim EtatCommunication As Integer       '�tat de la communication
    Dim AdresseFlottant As Long                'adresse du flottant
    Dim TValeurFlottant(1) As Single           'tableau contenant le flottant
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrFlottants = 1
    TValeurFlottant(1) = ValeurFlottant
    
    '--- d�termination de l'adresse ---
    AdresseFlottant = CLng(Right(LibelleFlottant, Len(LibelleFlottant) - 2))

    '--- �criture ---
    Call writefword(NumCanal, NumEquipement, NbrFlottants, AdresseFlottant, TValeurFlottant(1), EtatCommunication)
    
    '--- valeur de retour ---
    APIEcritureFlottant = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue une �criture d'un octet dans l'automate
' Entr�es :        LibelleOctet -> Libell� du bit (exemple : MB10)
'                        ValeurOctet -> Valeur de l'octet � �crire
' Retours : APILectureOctet -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APIEcritureOctet(ByVal LibelleOctet As String, _
                                                     ByVal ValeurOctet As Integer) As Integer
    
    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next
    'Call Log("APIEcritureOctet", "LibelleOctet:" & LibelleOctet & " ;ValeurOctet=" & ValeurOctet)
    '--- constantes priv�es ---
    
    '--- d�claration ---
    Dim NbrOctets As Integer                             'nombre d'octets
    Dim NumCanal As Integer                            'num�ro du canal
    Dim NumEquipement As Integer                  'num�ro de l'�quipement
    Dim EtatCommunication As Integer             '�tat de la communication
    Dim AdresseOctet As Long                          'adresse de l'octet
    Dim TValeurOctet(1) As Byte                        'tableau contenant la valeur de l'octet
    
    '--- affectation ---
    NumCanal = 0
    NumEquipement = 0
    NbrOctets = 1
    TValeurOctet(1) = ValeurOctet
    
    '--- d�termination de l'adresse ---
    AdresseOctet = CLng(Right(LibelleOctet, Len(LibelleOctet) - 2))
    
    '--- �criture ---
    Call writepackbyte(NumCanal, NumEquipement, NbrOctets, AdresseOctet, TValeurOctet(1), EtatCommunication)

    '--- valeur de retour ---
    APIEcritureOctet = EtatCommunication

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Effectue la condamnation ou d�condamnation d'un poste
' Entr�es :                                                 NumPoste -> Num�ro d'un poste
'                                                             EtatSouhaite -> FALSE = D�condamnation
'                                                                                        TRUE = Condamnation
' Retours : APICondamnationDecondamnationPoste -> 0 = Transmission bonne, sinon num�ro de l'erreur
' D�tails  :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function APICondamnationDecondamnationPoste(ByVal NumPoste As POSTES, _
                                                                                           ByVal EtatSouhaite As Boolean) As Long

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- constantes priv�es ---
    Const NOM_GROUPE = "SUIVI_LIGNE"    'nom du groupe
    
    '--- d�claration ---
    Dim NomVariable As String                       'nom de la variable OPC
        
    '--- affectation du nom de la variable ---
    NomVariable = "CondamnationPoste" & Right("00" & CStr(NumPoste), 2)
                    
    '--- �criture dans l'automate ---
    APICondamnationDecondamnationPoste = APIEcritureVariableNommee(NOM_GROUPE, NomVariable, IIf(EtatSouhaite = False, "0", "1"))

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
' R�le      : Affichage des donn�es transmises � l'API
' Entr�es :
' Retours :
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Sub AffichageDonneesTransmisesAPI(ByVal TexteAAfficher As String)

    '--- aiguillage en cas d'erreurs ---
    On Error Resume Next

    '--- transmission dans la Fenetre principale
    OccFPrincipale.LDonneesTransmisesAPI = TexteAAfficher

End Sub


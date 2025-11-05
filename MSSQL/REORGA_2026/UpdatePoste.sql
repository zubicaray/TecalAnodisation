USE [ANODISATION_2026]
-- 1. Création de la table UpdatePostes
-- J'utilise INT pour les deux colonnes car ce sont des nombres entiers.
IF OBJECT_ID('UpdatePostes', 'U') IS NOT NULL
    DROP TABLE UpdatePostes;

CREATE TABLE UpdatePostes (
    OldPoste INT NOT NULL,
    NewPoste INT NOT NULL,
    -- Ajouter une clé primaire si cette table est destinée à un usage fréquent de jointures.
    -- CONSTRAINT PK_UpdatePostes PRIMARY KEY (OldPoste)
);

-- 2. Insertion des données
-- Utilisation d'une seule commande INSERT avec de multiples lignes (VALUES) pour une meilleure performance.
INSERT INTO UpdatePostes (OldPoste, NewPoste)
VALUES
(1, 1),
(2, 2),
(4, 4),
(5, 3),
(6, 5),
(7, 6),
(8, 7),
(9, 8),
(10, 9),
(11, 10),
(12, 11),
(13, 12),
(14, 13),
(15, 14),
(16, 15),
(17, 16),
(18, 17),
(19, 18),
(20, 19),
(21, 20),
(22, 21),
(23, 22),
(24, 23),
(25, 24),
(26, 25),
(27, 26),
(28, 27),
(29, 28),
(30, 29),
(31, 30),
(32, 31),
(33, 32),
(34, 33),
(35, 34),
(36, 35),
(37, 36),
(38, 37),
(39, 38),
(40, 39),
(41, 44),
(42, 45),
(43, 41),
(44, 42);

-- 3. Vérification (Optionnel)
SELECT * FROM UpdatePostes ORDER BY OldPoste;
--****************************************************
--****************************************************
--****************************************************
IF OBJECT_ID('UpdateZones', 'U') IS NOT NULL
    DROP TABLE UpdateZones;

CREATE TABLE UpdateZones (
    OldZone INT NOT NULL,
    NewZone INT NOT NULL
    -- Vous pouvez ajouter une clé primaire sur OldZone si nécessaire: 
    -- CONSTRAINT PK_UpdateZones PRIMARY KEY (OldZone)
);

-- 2. Insertion des données
INSERT INTO UpdateZones (OldZone, NewZone)
VALUES
(39, 3),
(3, 4),
(4, 5),
(5, 6),
(6, 7),
(7, 8),
(8, 9),
(9, 10),
(10, 11),
(11, 12),
(12, 13),
(13, 14),
(14, 15),
(15, 16),
(16, 18),
(17, 19),
(18, 20),
(19, 21),
(20, 22),
(21, 23),
(22, 24),
(23, 25),
(24, 26),
(25, 27),
(26, 28),
(27, 29),
(28, 30),
(29, 31),
(30, 32),
(31, 33),
(32, 34),
(33, 35),
(36, 36), -- Zone inchangée
(34, 37),
(35, 42),
(37, 39),
(38, 40);

-- 3. Vérification (Optionnel)
-- Affiche le contenu de la nouvelle table
SELECT OldZone, NewZone 
FROM UpdateZones 
ORDER BY OldZone;


--****************************************************
--****************************************************
--****************************************************


UPDATE F
SET 
    F.NumPoste = U1.NewPoste,
    F.NumPostePrecedent = U2.NewPoste
FROM 
    DetailsFichesProduction AS F
INNER JOIN 
    UpdatePostes AS U1 ON U1.OldPoste = F.NumPoste
INNER JOIN 
    UpdatePostes AS U2 ON U2.OldPoste = F.NumPostePrecedent;


--****************************************************
--****************************************************
--****************************************************


UPDATE T
SET 
    T.NumPosteDepart = U1.NewPoste,
    T.NumPosteArrivee= U2.NewPoste
FROM 
    TempsMouvementsTranslationPonts AS T
INNER JOIN 
    UpdatePostes AS U1 ON U1.OldPoste = T.NumPosteDepart
INNER JOIN 
    UpdatePostes AS U2 ON U2.OldPoste = T.NumPosteArrivee;

--****************************************************
--****************************************************
--****************************************************


UPDATE T
SET 
    T.Depart = U1.NewZone,
    T.Arrivee= U2.NewZone
FROM 
    TempsDeplacements AS T
INNER JOIN 
    UpdateZones AS U1 ON U1.OldZone = T.Depart
INNER JOIN 
    UpdateZones AS U2 ON U2.OldZone = T.Arrivee;

--****************************************************
--****************************************************
--****************************************************


UPDATE G
SET 
    G.NumZone = Z.NewZone,
    G.NumPosteReel = U.NewPoste
FROM 
    DetailsGammesProduction AS G
INNER JOIN 
    UpdateZones AS Z ON Z.OldZone = G.NumZone
INNER JOIN 
    UpdatePostes AS U ON U.OldPoste = G.NumPosteReel;

--****************************************************
--****************************************************
--****************************************************


UPDATE G
SET 
    G.NumZone = Z.NewZone
FROM 
    DetailsGammesAnodisation AS G
INNER JOIN 
    UpdateZones AS Z ON Z.OldZone = G.NumZone;



--****************************************************
--****************************************************
--****************************************************


UPDATE P
SET 
    P.NumPosteDepart = U1.NewPoste,
    P.NumPosteArrivee = U2.NewPoste
FROM 
    PREMISSES AS P
INNER JOIN 
    UpdatePostes AS U1 ON U1.OldPoste = P.NumPosteDepart
INNER JOIN 
    UpdatePostes AS U2 ON U2.OldPoste = P.NumPosteArrivee;



--****************************************************
--****************************************************
--****************************************************

INSERT INTO PREMISSES
    (NumPont, NumPontIA, NumPosteDepart, NumPosteArrivee, PremisseCodee, PremisseDecodee, 
    TempsCycleSecondes)
SELECT
    2, 
    2,
    P_A.NumPoste,
    P_B.NumPoste,
    '', 
    '', 
    0
FROM
    POSTES AS P_A
    CROSS JOIN POSTES AS P_B -- Utilisation explicite du CROSS JOIN, plus clair que la virgule
WHERE NOT EXISTS (
    SELECT 1 
    FROM PREMISSES AS Pr
    WHERE Pr.NumPosteDepart = P_A.NumPoste
      AND Pr.NumPosteArrivee = P_B.NumPoste
);
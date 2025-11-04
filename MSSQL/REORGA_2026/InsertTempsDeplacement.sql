

WITH AllCombinations AS (
    SELECT
        t1.n AS depart,
        t2.n AS arrivee
    FROM (SELECT TOP 46 ROW_NUMBER() OVER(ORDER BY (SELECT NULL)) AS n FROM sys.objects t1, sys.objects t2) AS t1 -- Génère une série de nombres de 1 à 85
    JOIN (SELECT TOP 46 ROW_NUMBER() OVER(ORDER BY (SELECT NULL)) AS n FROM sys.objects t1, sys.objects t2) AS t2 ON 1 = 1 -- Génère une autre série de nombres de 1 à 85
)
INSERT INTO TempsDeplacements (depart, arrivee, lent, normal, rapide)
-- Sélectionne uniquement les couples qui n'existent pas encore dans la table.
SELECT
    ac.depart,
    ac.arrivee,
    0 AS lent,   -- Les valeurs par défaut sont 0
    0 AS normal,
    0 AS rapide
FROM
    AllCombinations ac
LEFT JOIN
    TempsDeplacements td ON ac.depart = td.depart AND ac.arrivee = td.arrivee
WHERE
    td.depart IS NULL;
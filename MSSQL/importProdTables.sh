#!/bin/bash

# Définissez les paramètres de connexion
SERVER="localhost"   # Remplacez par votre serveur SQL
DATABASE="ANODISATION"
USER="sa"  # Remplacez par votre utilisateur SQL
PASSWORD="Jeff_nenette"  # Remplacez par votre mot de passe SQL

# Ajout de l'option TrustServerCertificate=yes pour éviter l'erreur SSL
OPTIONS="-S $SERVER -d $DATABASE -U $USER -P $PASSWORD -C -l 6000"

# Récupérer le répertoire du script
SCRIPT_DIR="$(dirname "$0")"

# Boucle pour chaque fichier SQL dans le répertoire du script
for sql_file in "$SCRIPT_DIR"/export*.sql; do
    if [ -f "$sql_file" ]; then
        echo "Traitement du fichier SQL : $sql_file"

        sqlcmd $OPTIONS -i "$sql_file"
    else
        echo "Aucun fichier SQL trouvé dans le dossier $SCRIPT_DIR."
    fi
done

echo "Tous les scripts SQL dans le dossier ont été exécutés."

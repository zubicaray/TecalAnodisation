#!/bin/bash

# Définissez les paramètres de connexion
SERVER="localhost"   # Remplacez par votre serveur SQL
DATABASE="ANODISATION"
USER="sa"  # Remplacez par votre utilisateur SQL
PASSWORD="Jeff_nenette"  # Remplacez par votre mot de passe SQL
#!/bin/bash


# Ajout de l'option TrustServerCertificate=yes pour éviter l'erreur SSL
OPTIONS="-S $SERVER -d $DATABASE -U $USER -P $PASSWORD -C "

# Commande pour obtenir la liste des tables
tables=$(sqlcmd $OPTIONS -Q "SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE';" -h -1)

# Vérifier s'il y a des tables à supprimer
if [ -z "$tables" ]; then
    echo "Aucune table trouvée dans la base de données $DATABASE."
    exit 0
fi

# Boucle pour supprimer chaque table
for table in $tables
do
    echo "Suppression de la table: $table"
    sqlcmd $OPTIONS -Q "DROP TABLE [$table];"
done

echo "Toutes les tables ont été supprimées de la base de données $DATABASE."

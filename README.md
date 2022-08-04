# CRCRME - Générer répertoire métiers

Ce répertoire sert à générer un fichier JSON contenant toutes les informations nécessaires sur les différents métiers semi-spécialisé selon le [Ministère de l'Éducation du Québec](http://www1.education.gouv.qc.ca/sections/metiers/index.asp).

## Usage

1. Installer python
2. Démarrer [genererRepertoireMetiers.py](genererRepertoireMetiers.py).
3. Choisir le fichier Excel contenant les informations sur les métiers.
4. Cliquer sur générer et attendre que le programme affiche `Tout est fini!`.

## Modification du code

Si de nouvelles colones sont ajoutée au fichier Excel, il est nécessaire d'ajouter un champ dans le dictionnaire `EXCEL_DATA_HEADERS`.  
La `key` correspond au nom du champ dans le JSON généré.  
La `value` correspond au nom exact de la colone dans Excel.

# WordStylesIndexor

 Script pour l'extraction de données par les styles dans des documents Word

# Présentation

WordStyleIndexor est un script que j’ai créé pour démontrer, dans le cadre du cours [INU1010 - Création de l'information numérique]( https://cours.ebsi.umontreal.ca/planscours/inu1010), l’extraction de données depuis un document Word structuré à l’aide d’une feuille de style.

Plus précisément :

- Ce script extrait dans une collection de documents Word placés dans le même répertoire que lui, les contenus structurés par styles.

- La liste des contenus « stylés » à extraire doit être spécifiée à l'aide de la variable *stylesAIndexer*.
  Par défaut, celui-ci va extraire les informations possédant les styles « Titre 1 » et « Titre 2 »

- Les contenus sont extraits vers un document CSV qui contient : un identifiant, le nom du fichier et le nom du style.

- Une base de données Access est également créée : pour faire des recherches, ouvrez cette base de données et utilisez alors le mode « requête ».

Cet outil, en VBS, est lui-même un assemblage de plusieurs morceaux de code que j’ai trouvé sur Internet et dont la provenance est habituellement indiquée dans le code source.

# Utilisation

* Placez les fichiers *WordStylesIndexor.vbs* et *go.bat* dans le même dossier que les documents à indexer.
* Au besoin, rendez-vous tout à tour dans les propriétés des fichiers *WordStylesIndexor.vbs* et *go.bat* et cliquez sur le bouton *Débloquer*.
* Éditez le fichier *WordStylesIndexor.vbs* avec un éditeur texte de votre choix pour éditer la variable *stylesAIndexer*.
* Double-cliquez sur le fichier *go.bat* pour exécuter le script.

**Création automatique de la base de données Access**

Si vous utilisez Office 365 sur votre poste de travail, il est possible que vous deviez installer [Microsoft Access Database Engine 2016 (x86)](https://www.microsoft.com/fr-FR/download/details.aspx?id=54920) pour que le script puisse créer automatiquement la base de données Access.
Comme mentionné dans la [documentation de Microsoft]( https://docs.microsoft.com/en-us/office/troubleshoot/access/cannot-use-odbc-or-oledb), dans un tel cas Microsoft Access Database Engine 2016 doit être installé en mode « quiet ».
Pour vous assister dans cette étape :

* Téléchargez le fichier d’installation accessdatabaseengine.exe de [Microsoft Access Database Engine 2016 (x86)](https://www.microsoft.com/fr-FR/download/details.aspx?id=54920) depuis le site de Microsoft.
* Placez le fichier d’installation dans le dossier MADE2006.
* Exécutez le fichier install.bat en tant qu’administrateur.

# Limitations

* Ne fonctionne que sur Windows.
* Les fichiers .docx à indexer doivent être dans le même dossier que le script.
* Ne fonctionne pas sur les documents verrouillés.
* Ne fonctionne pas avec les cases à cocher.
* Est instable, surtout si les documents sont déjà ouverts dans Word, ou vous tentez d’indexer des styles qui n’existent pas dans les documents.

# Licence

Cette création est mise à disposition selon le Contrat Paternité-NonCommercial-ShareAlike2.5 Canada disponible en ligne http://creativecommons.org/licenses/by-nc-sa/2.5/ca/ ou par courrier postal à Creative Commons, 559 Nathan Abbott Way, Stanford, California 94305, USA.

# Changelog

**2009-09-23**

- Première version disponible

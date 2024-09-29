# WordStylesIndexor

 Script pour l'extraction de données par les styles dans des documents Word

# Présentation

WordStyleIndexor est un script que j’ai créé pour démontrer, dans le cadre du cours [INU1010 - Création de l'information numérique]( https://cours.ebsi.umontreal.ca/planscours/inu1010), l’extraction de données depuis un document Word structuré à l’aide d’une feuille de style.

Plus précisément :
* Ce script extrait dans une collection de documents Word placés dans le même répertoire que lui, les contenus structurés par styles.
* La liste des contenus « stylés » à extraire doit être spécifiée à l'aide de la variable *stylesAIndexer*.  
Par défaut, celui-ci va extraire les informations possédant les styles « Titre 1 » et « Titre 2 »
* Les contenus sont extraits vers un document CSV qui contient : un identifiant, le nom du fichier et le nom du style.
* Une base de données Access est également créée : pour faire des recherches, ouvrez cette base de données et utilisez alors le mode « requête ».

Cet outil, en VBS, est lui-même un assemblage de plusieurs morceaux de code que j’ai trouvé sur Internet et dont la provenance est habituellement indiquée dans le code source.

# Utilisation

* Placez les fichiers *WordStylesIndexor.vbs* et *go.bat* dans le même dossier que les documents à indexer.
* Au besoin, rendez-vous tout à tour dans les propriétés des fichiers *WordStylesIndexor.vbs* et *go.bat* et cliquez sur le bouton *Débloquer*.
* Éditez le fichier *WordStylesIndexor.vbs* avec un éditeur texte de votre choix pour éditer la variable *stylesAIndexer*.
* Double-cliquez sur le fichier *go.bat* pour exécuter le script.

# Limitations

* Ne fonctionne que sur Windows.
* Les fichiers .docx à indexer doivent être dans le même dossier que le script.
* Ne fonctionne pas sur les documents verrouillés.
* Ne fonctionne pas avec les cases à cocher.
* Est instable, surtout si les documents sont déjà ouverts dans Word, ou vous tentez d’indexer des styles qui n’existent pas dans les documents.

# Licence

Cette création est mise à disposition selon le Contrat Paternité-NonCommercial-ShareAlike2.5 Canada disponible en ligne http://creativecommons.org/licenses/by-nc-sa/2.5/ca/ ou par courrier postal à Creative Commons, 559 Nathan Abbott Way, Stanford, California 94305, USA.

# Changelog

**2022-09-29**
* Meilleur support des systèmes 64 bits
* Le script fonctionne avec ou sans utilisation du fichier *go.bat* (en double-cliquant sur *WordStylesIndexor.vbs*).
* Révision de la documentation : l'installation de Microsoft Access Database Engine 2016 n'est plus requise avec O365 ou Office 2021 et plus.

**2020-11-03**
* Provider Jet remplacé par OLEDB, pour création d'une BD .accdb au lieu d'une BD .mdb
* Ne pas créer la BD si ADOX n'est pas disponible (Office 365)

**2013-10-28**
* Ajout *go.bat*

**2012-11-11**
* Compatible avec documents Office 2010
* Ne fonctionne pas sur les systèmes 64 bits

**2009-09-23**
* Première version disponible

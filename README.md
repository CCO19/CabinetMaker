# CabinetMaker

Powershell script generating a **GUI** to generate Cabinet archives via **MakeCAB** in a semi-automated manner.

L'interface graphique permet :

- Ajouter des fichiers et/ou des dossiers de fichiers à l'achive Cabinet _(en préservant ou non l'aborescence)_
- Activation/Désactivation de la compression
- Choix du type de compression _(**LZX** ou **MSZIP**)_
- Niveau de compression _(de **15** à **21**)_
- Activation/Désactivation du découpage de l'archive Cabinet en volumes de taille choisie _(par défaut de **1 Mo** à **2Go**)_
- Choisir le répertoire de sortie de l'archive Cabinet générée

Le script est utilisable mais en toujours en version bêta. Il peut donc présenter des bugs, notement dans la gestion des caractères spéciaux.

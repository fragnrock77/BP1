# Explorateur de données CSV & XLSX

Application web monopage permettant d'importer des fichiers CSV ou XLSX volumineux, de rechercher plusieurs mots-clés (avec opérateurs logiques) et d'exporter les résultats filtrés.

## Fonctionnalités clés
- Importation par glisser-déposer ou sélection de fichier avec barre de progression.
- Analyse côté client optimisée par flux pour les CSV et lecture par blocs pour les XLSX.
- Recherche multi-mots-clés sensible/insensible à la casse, exacte ou partielle, exécutée dans un Web Worker.
- Tableau paginé avec tri par colonne et options d'export (presse-papiers, CSV, XLSX).

## Utilisation
1. Ouvrir `index.html` dans un navigateur moderne (Chrome, Edge, Firefox, Safari).
2. Déposer un fichier CSV ou XLSX (≤ 600 Mo) ou utiliser le sélecteur de fichier.
3. Une fois les données chargées, saisir une requête de recherche en utilisant `AND`, `OR`, `NOT` ou `-mot`.
4. Parcourir les résultats via la pagination et exporter selon le format souhaité.

> Les opérations d'analyse de fichier et de recherche sont exécutées hors du thread principal pour conserver une interface réactive même avec de gros volumes de données.

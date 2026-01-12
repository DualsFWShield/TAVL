# Relevés TAVL

Outil web pour l'importation, l'édition et l'exportation de fichiers de relevés XLSX (structure Matrix).

## Fonctionnalités

- **Importation Excel** : Chargez vos fichiers `.xlsx` (ex: `Barbe.xlsx`). L'outil analyse automatiquement la structure du fichier (Catégories, Questions, Types de données).
- **Interface Premium** : Une interface utilisateur sombre et moderne (Dark Mode) pour une saisie de données agréable.
- **Formulaires Dynamiques** : Génération automatique des champs de saisie basée sur les en-têtes du fichier Excel :
  - **Vrai/Faux (v/f)** et **Oui/Non (o/n)** : Boutons radio intuitifs.
  - **Gradin/Mobile/Fixe (G/M/F)** : Sélecteur rapide à trois états.
  - **Dates** : Sélecteur de date (calendrier).
  - **Nombres** : Champs numériques.
  - **Texte** : Zones de texte auto-extensibles.
- **Navigation Latérale** : Liste des "Auditoires" générée dynamiquement pour une navigation rapide entre les locaux.
- **Sauvegarde Automatique** : Vos modifications sont enregistrées localement dans votre navigateur en temps réel. En cas de fermeture accidentelle, vous pouvez restaurer votre session.
- **Exportation Intelligente** : Exportez le fichier modifié au format XLSX en **conservant le formatage original** (couleurs, polices, bordures) du fichier source.

## Utilisation

1. Ouvrez le fichier `index.html` dans un navigateur web moderne (Chrome, Edge, Firefox).
2. Glissez-déposez votre fichier Excel (ex: `Barbe.xlsx`) dans la zone dédiée.
3. Sélectionnez un auditoire dans la barre latérale gauche.
4. Remplissez ou modifiez les informations dans le formulaire central.
5. Vos changements sont sauvegardés automatiquement.
6. Cliquez sur **"Exporter le relevé"** en haut à droite pour télécharger le fichier Excel mis à jour.

## Structure du Fichier Excel Attendue

L'outil s'attend à une structure de type "Matrice" spécifique :
- **Ligne 3** : Catégories (Header principal)
- **Ligne 4** : Questions / Libellés (Header secondaire)
- **Ligne 5** : Types de données (`v/f`, `o/n`, `nombre`, `date`, `text`, etc.)
- **Ligne 6 et suivantes** : Données (Une ligne par Auditoire)
- **Colonne "Auditoires"** : Utilisée pour identifier les locaux dans la barre latérale.

## Technologies

- **HTML5 / CSS3** (Vanilla, sans framework lourd)
- **JavaScript** (ES6+)
- **ExcelJS** : Bibliothèque pour la manipulation avancée des fichiers Excel dans le navigateur.
- **IndexedDB** : Pour la persistance des données locale.
- **Google Fonts** : Police "Inter".

## Installation pour Développement

Aucune installation complexe requise (pas de Node.js ou npm obligatoire pour l'exécution).
Il suffit de servir les fichiers via un serveur local ou d'ouvrir `index.html` directement.

Pour cloner le projet :
```bash
git clone <votre-depot>
```

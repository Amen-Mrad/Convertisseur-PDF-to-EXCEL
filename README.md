# Convertisseur Attijari Bank PDF vers Excel

## Description
Application Python pour convertir automatiquement les extraits de compte PDF d'Attijari Bank en fichiers Excel avec les colonnes : Date, Libellé, Débit, Crédit.

## Fonctionnalités
- Interface graphique simple et intuitive
- Détection automatique des fichiers Attijari Bank
- Extraction des 4 colonnes principales : Date, Libellé, Débit, Crédit
- Sauvegarde automatique sur le bureau
- Gestion des formats de date DD/MM/YYYY
- Gestion des montants avec virgule comme séparateur décimal

## Installation
1. Installer Python 3.7 ou plus récent
2. Installer les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation
1. Lancer l'application :
```bash
python attijari_converter.py
```

2. Sélectionner le fichier PDF Attijari Bank
3. Entrer le nom souhaité pour le fichier Excel
4. Cliquer sur "Convertir PDF vers Excel"
5. Le fichier Excel sera créé automatiquement sur le bureau

## Format des données extraites
- **Date** : Format DD/MM/YYYY
- **Libellé** : Description de la transaction
- **Débit** : Montant débité (vide si crédit)
- **Crédit** : Montant crédité (vide si débit)

## Compatibilité
- Compatible avec tous les extraits Attijari Bank
- Gère les PDFs textuels (pas d'OCR nécessaire)
- Format Excel : .xlsx

## Notes techniques
- Utilise pdfplumber pour l'extraction de données
- Détection automatique des tableaux
- Gestion des libellés multi-lignes
- Format des montants : virgule comme séparateur décimal

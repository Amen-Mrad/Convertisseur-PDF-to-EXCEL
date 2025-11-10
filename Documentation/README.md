# Convertisseur Multi-Banques PDF vers Excel

## Description
Application Python universelle pour convertir automatiquement les extraits de compte PDF de différentes banques tunisiennes en fichiers Excel avec les colonnes : Date, Libellé, Débit, Crédit.

## Banques supportées
- **Attijari Bank** (extraits et relevés)
- **Amen Bank** (extraits et relevés)
- **BIAT** (extraits et relevés)
- **BNA** (extraits et relevés)
- **BT** (extraits et relevés)
- **BTK** (extraits et relevés)
- **QNB** (relevés)
- **STB** (extraits et relevés)
- **UBCI** (extraits)
- **UIB** (relevés)
- **Zitouna Bank** (extraits et relevés)

## Fonctionnalités
- Interface graphique simple et intuitive
- Détection automatique du type de banque
- Support des extraits et relevés de compte
- Extraction des 4 colonnes principales : Date, Libellé, Débit, Crédit
- Sauvegarde automatique sur le bureau
- Gestion des formats de date DD/MM/YYYY
- Gestion des montants avec virgule comme séparateur décimal
- Support OCR pour les PDFs scannés (BNA)

## Installation
1. Installer Python 3.7 ou plus récent
2. Installer les dépendances :
```bash
pip install -r requirements.txt
```

## Utilisation
1. Lancer l'application principale :
```bash
python lancer_convertisseur.py
```

2. Sélectionner le fichier PDF de votre banque
3. Entrer le nom souhaité pour le fichier Excel
4. Cliquer sur "Convertir PDF vers Excel"
5. Le fichier Excel sera créé automatiquement sur le bureau

## Format des données extraites
- **Date** : Format DD/MM/YYYY
- **Libellé** : Description de la transaction
- **Débit** : Montant débité (vide si crédit)
- **Crédit** : Montant crédité (vide si débit)

## Compatibilité
- Compatible avec tous les extraits et relevés des banques supportées
- Gère les PDFs textuels et scannés (OCR pour BNA)
- Format Excel : .xlsx

## Notes techniques
- Utilise pdfplumber pour l'extraction de données
- Détection automatique des tableaux
- Gestion des libellés multi-lignes
- Format des montants : virgule comme séparateur décimal
- Support OCR avec Tesseract pour les PDFs scannés

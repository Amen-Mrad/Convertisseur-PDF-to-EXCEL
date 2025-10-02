# Convertisseur Attijari Bank - Structure Complète

## 📁 Structure du Projet

```
convertisseur_bancaire/
├── attijari_converter.py          # Convertisseur EXTRAT DE COMPTE Attijari
├── attijari_releve_converter.py   # Convertisseur RELEVE DE COMPTE Attijari
├── biat_releve_converter.py       # Convertisseur RELEVE DE COMPTE BIAT
├── lancer_convertisseur.py        # Script de lancement principal
├── requirements.txt               # Dépendances Python
├── logo/                         # Dossier pour les logos
│   ├── attijari.png              # Logo Attijari Bank
│   └── biat.png                  # Logo BIAT
└── README_STRUCTURE.md           # Ce fichier
```

## 🚀 Utilisation

### Script de lancement principal :
```bash
python lancer_convertisseur.py
```

### Ou directement :
```bash
python attijari_converter.py          # EXTRAT Attijari
python attijari_releve_converter.py   # RELEVE Attijari
python biat_releve_converter.py       # RELEVE BIAT
```

## 📋 Fonctionnalités

### ✅ EXTRAT DE COMPTE (`attijari_converter.py`)
- Interface verte
- Détection automatique des tableaux
- Parsing des colonnes : Date, Libellé, Débit, Crédit
- Formatage Excel avec fond jaune et bordures
- Nom de feuille : "J03"

### ✅ RELEVE DE COMPTE Attijari (`attijari_releve_converter.py`)
- Interface rouge
- Détection d'année dynamique depuis le PDF
- Parsing adapté pour structure RELEVE (6 colonnes)
- Même formatage Excel que EXTRAT
- Support des formats de date : DD MM, DD MM YYYY

### ✅ RELEVE DE COMPTE BIAT (`biat_releve_converter.py`)
- Interface violette
- Détection spécifique BIAT (mots-clés, logo)
- Parsing adapté pour structure BIAT (5 colonnes)
- Même formatage Excel que les autres
- Support des formats de date : DD MM, DDMMYYYY

## 🎯 Formats Supportés

### EXTRAT :
- Structure : Date, Libellé, Débit, Crédit, Solde
- Format date : DD/MM/YYYY

### RELEVE Attijari :
- Structure : DATE OPERATION, LIBELLE, DATE VALEUR, DEBIT (TND), CREDIT (TND), SOLDE (TND)
- Formats date : DD MM, DD MM YYYY
- Détection automatique de l'année

### RELEVE BIAT :
- Structure : Date, Libellé, Date de valeur, Débit, Crédit
- Formats date : DD MM, DDMMYYYY
- Détection automatique de l'année

## 📊 Sortie Excel

Les deux convertisseurs génèrent des fichiers Excel avec :
- **En-têtes** : Fond jaune, police Arial 12pt gras
- **Feuille** : Nommée "J03"
- **Bordures** : Toutes les cellules
- **Colonnes** : Date, Libellé, Débit, Crédit
- **Largeurs** : Date (12), Libellé (50), Débit/Crédit (15)
- **Sauvegarde** : Sur le bureau de l'utilisateur

## 🔧 Installation

```bash
pip install -r requirements.txt
```

## 📝 Dépendances

- `tkinter` (inclus avec Python)
- `pdfplumber` (extraction PDF)
- `pandas` (manipulation données)
- `openpyxl` (formatage Excel)
- `PIL` (traitement images)
- `PyMuPDF` (extraction images PDF)
- `opencv-python` (détection logo)
- `numpy` (traitement images)

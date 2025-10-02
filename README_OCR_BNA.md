# 🔍 Support OCR pour les PDFs BNA Scannés

## Fonctionnalités Ajoutées

### 1. **Détection Automatique des PDFs Scannés**
- Le convertisseur détecte automatiquement si un PDF est scanné (pas de texte extractible)
- Bascule automatiquement vers l'OCR si nécessaire

### 2. **OCR Optimisé pour les Relevés Bancaires**
- **Pré-traitement avancé** : Réduction du bruit, amélioration du contraste
- **Résolution élevée** : Matrix 3x3 pour une meilleure qualité
- **Configuration spécialisée** : Optimisée pour les documents bancaires
- **Nettoyage intelligent** : Correction des erreurs OCR communes

### 3. **Interface Utilisateur Améliorée**
- Messages informatifs pendant le traitement OCR
- Indication claire quand l'OCR est utilisé
- Barre de progression mise à jour

### 4. **Techniques de Pré-traitement**
- Conversion en niveaux de gris
- Réduction du bruit avec `medianBlur`
- Amélioration du contraste avec `CLAHE`
- Seuillage adaptatif `OTSU`
- Morphologie pour nettoyer les caractères

### 5. **Nettoyage du Texte OCR**
- Correction des erreurs communes (O → 0)
- Filtrage des lignes non-pertinentes
- Normalisation des espaces
- Détection des données bancaires

## Utilisation

1. **Lancer le convertisseur BNA** : `python bna_releve_converter.py`
2. **Sélectionner un PDF scanné** : Le système détectera automatiquement le type
3. **Traitement automatique** : L'OCR se lance si nécessaire
4. **Résultat** : Fichier Excel avec toutes les transactions extraites

## Dépendances Requises

```bash
pip install PyMuPDF pytesseract Pillow numpy opencv-python
```

## Installation Tesseract

- **Windows** : Télécharger depuis [UB-Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
- **Linux** : `sudo apt-get install tesseract-ocr tesseract-ocr-fra`
- **macOS** : `brew install tesseract tesseract-lang`

## Avantages

✅ **Traitement automatique** des PDFs scannés  
✅ **Qualité OCR optimisée** pour les relevés bancaires  
✅ **Interface utilisateur claire** avec feedback  
✅ **Robustesse** avec gestion d'erreurs  
✅ **Performance** avec pré-traitement intelligent

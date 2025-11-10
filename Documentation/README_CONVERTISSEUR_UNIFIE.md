# 🚀 CONVERTISSEUR UNIFIÉ MULTI-BANQUES

## 🎯 **Vue d'Ensemble**

Le **Convertisseur Unifié** est une interface unique qui détecte automatiquement le type de banque et de document (Extrait/Relevé) à partir d'un fichier PDF, puis convertit automatiquement vers Excel sans intervention manuelle.

## ✨ **Fonctionnalités Principales**

### 🔍 **Détection Automatique**
- **Reconnaissance de banque** : AMEN, ZITOUNA, BNA, STB, WIFAK, BIAT, BT, UBCI, QNB
- **Identification du type** : Extrait de compte ou Relevé de compte
- **Validation automatique** : Vérification du format et des données

### 🎮 **Interface Simplifiée**
- **Un seul fichier** : Plus besoin de choisir manuellement
- **Drag & Drop** : Sélection simple du PDF
- **Feedback visuel** : Affichage de la détection en temps réel
- **Conversion en un clic** : Bouton unique pour tout le processus

### 📊 **Résultats Professionnels**
- **Format Excel standardisé** : Colonnes Date, Libellé, Débit, Crédit
- **Formatage automatique** : Styles, couleurs, bordures
- **Sauvegarde intelligente** : Fichier dans Downloads/ avec timestamp

---

## 🚀 **Utilisation**

### **Installation**
```bash
# 1. Installer les dépendances
pip install -r requirements.txt

# 2. Lancer le convertisseur unifié
python convertisseur_unifie.py
```

### **Workflow Utilisateur**
1. **📁 Sélectionner** le fichier PDF bancaire
2. **🔍 Détection automatique** de la banque et du type
3. **📝 Vérifier** les informations détectées
4. **🚀 Cliquer** sur "CONVERTIR AUTOMATIQUEMENT"
5. **✅ Récupérer** le fichier Excel dans Downloads/

---

## 🏦 **Banques Supportées**

| Banque | Extrait | Relevé | Statut |
|--------|---------|--------|--------|
| **AMEN BANK** | ✅ | ✅ | Actif |
| **ZITOUNA BANK** | ✅ | ✅ | Actif |
| **BNA BANK** | ✅ | ✅ | Actif |
| **STB BANK** | ✅ | ✅ | Actif |
| **WIFAK BANK** | ✅ | ✅ | Actif |
| **BIAT BANK** | ❌ | ✅ | Actif |
| **BT BANK** | ❌ | ✅ | Actif |
| **UBCI BANK** | ✅ | ❌ | Actif |
| **QNB BANK** | ❌ | ✅ | En cours |

---

## 🔧 **Architecture Technique**

### **Composants Principaux**
- **Interface Unifiée** : `ConvertisseurUnifie` - Classe principale
- **Détection Intelligente** : `detecter_banque_et_type()` - Reconnaissance automatique
- **Conversion Dynamique** : `appeler_convertisseur_specifique()` - Appel des convertisseurs spécialisés
- **Fallback Générique** : `convertir_generique()` - Pour les fichiers non reconnus

### **Processus de Détection**
```
PDF Input → Extraction Texte → Analyse Patterns → Détection Banque/Type → Conversion Spécialisée
```

### **Mapping des Convertisseurs**
```python
converter_mapping = {
    ('AMEN BANK', 'Extrait'): 'amen_extrait_converter.py',
    ('AMEN BANK', 'Relevé'): 'amen_releve_converter.py',
    ('ZITOUNA BANK', 'Extrait'): 'zitouna_extrait_converter.py',
    # ... etc
}
```

---

## 🧪 **Test et Validation**

### **Script de Test**
```bash
# Lancer le script de test
python test_convertisseur_unifie.py
```

### **Fonctionnalités Testées**
- ✅ Détection automatique des banques
- ✅ Identification des types de documents
- ✅ Appel des convertisseurs spécialisés
- ✅ Gestion des erreurs
- ✅ Interface utilisateur

---

## 🎯 **Avantages vs Interface Classique**

### **Interface Classique (lancer_convertisseur.py)**
- ❌ Sélection manuelle de la banque
- ❌ Choix manuel du type de document
- ❌ 3 étapes : Banque → Type → PDF
- ❌ Risque d'erreur de sélection

### **Interface Unifiée (convertisseur_unifie.py)**
- ✅ Détection automatique de la banque
- ✅ Identification automatique du type
- ✅ 1 étape : PDF → Excel
- ✅ Zéro risque d'erreur de sélection

---

## 🔮 **Évolutions Futures**

### **Améliorations Prévues**
- **Machine Learning** : Détection encore plus précise
- **Batch Processing** : Traitement multiple de fichiers
- **API REST** : Service web pour intégration
- **Cloud Storage** : Sauvegarde automatique

### **Nouvelles Fonctionnalités**
- **Prévisualisation** : Aperçu des données avant conversion
- **Validation** : Vérification des montants et dates
- **Historique** : Suivi des conversions précédentes
- **Export** : Formats multiples (CSV, JSON)

---

## 📋 **Cas d'Usage**

### **Particuliers**
- Conversion des relevés bancaires personnels
- Suivi des dépenses et revenus
- Archivage numérique des documents

### **Professionnels**
- Comptables : Traitement des documents clients
- Gestionnaires : Suivi des comptes d'entreprise
- Consultants : Analyse financière

### **Entreprises**
- Automatisation des processus comptables
- Intégration dans les systèmes existants
- Réduction des erreurs manuelles

---

## 🛠️ **Dépannage**

### **Problèmes Courants**

#### **Banque Non Détectée**
- Vérifier que le PDF contient du texte (pas seulement des images)
- Essayer la conversion générique
- Vérifier le format du PDF

#### **Erreur de Conversion**
- Vérifier que le convertisseur spécialisé existe
- Contrôler les permissions d'écriture dans Downloads/
- Vérifier la validité du fichier PDF

#### **Interface Ne Se Lance Pas**
- Vérifier l'installation des dépendances
- Contrôler la version Python (3.7+)
- Vérifier les permissions d'exécution

---

## 📞 **Support**

### **Logs de Debug**
Le convertisseur affiche des informations détaillées dans la console pour diagnostiquer les problèmes.

### **Fichiers de Test**
Utilisez `test_convertisseur_unifie.py` pour valider le fonctionnement.

### **Documentation**
Consultez `ANALYSE_PROJET.md` pour une vue d'ensemble complète du projet.

---

## 🎉 **Conclusion**

Le **Convertisseur Unifié** représente l'évolution naturelle du projet, offrant une expérience utilisateur simplifiée et une efficacité maximale. Plus besoin de connaître le type de document ou la banque - il suffit de glisser-déposer le PDF et la conversion se fait automatiquement !

**🚀 L'avenir de la conversion bancaire est là !**

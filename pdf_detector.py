#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Module de détection automatique des types de documents PDF bancaires
"""

import pdfplumber
import re
from typing import Dict, Optional, Tuple

class PDFBankDetector:
    """Détecteur automatique de types de documents PDF bancaires"""
    
    def __init__(self):
        # Mots-clés pour détecter les extraits Zitouna
        self.zitouna_extrait_keywords = [
            r'extrait\s+du\s+compte',
            r'extrait\s+de\s+compte',
            r'liste\s+des\s+transactions',
            r'banque\s+zitouna',
            r'zitouna\s+bank',
            r'mصرف\s+الزيتونة',  # Zitouna en arabe
        ]
        
        # Mots-clés pour détecter les extraits Amen Bank
        self.amen_extrait_keywords = [
            r'@mennet',
            r'@menet',
            r'amen\s+bank',
            r'بنك\s+الأمان',  # Amen Bank en arabe
            r'extrait\s+de\s+compte.*amen',
            r'liste\s+des\s+transactions.*amen',
        ]

        # Mots-clés pour détecter BTK (relevé/extrait)
        self.btk_keywords = [
            r'btk@direct',
            r'btk\s*bank',
            r'bank\s*tunisienne\s*korian',
            r'extrait\s+de\s+compte.*btk',
            r'virement\s+emis\s+aut\s+bq',
        ]
        
        # Mots-clés pour détecter BNA (relevé/extrait)
        self.bna_keywords = [
            r'banque\s+nationale\s+agricole',
            r'bna\s*bank',
            r'extrait\s+du\s+compte.*bna',
            r'relevé\s+de\s+compte.*bna',
            r'tunis\s+le\s*:',
        ]
        
        # Mots-clés pour détecter BT (Banque de Tunisie)
        self.bt_keywords = [
            r'banque\s+de\s+tunisie',
            r'البنك\s+التونسيي',  # Banque de Tunisie en arabe
            r'btbktntt',  # Code BIC
            r'relevé\s+mensuel',
            r'extrait\s+de\s+compte.*bt',
            r'compte.*tnd',
            r'rib.*0510',
            r'la\s+banque\s+de\s+tunisie',
            r'société\s+anonyme.*capital.*225',
            r'rue\s+de\s+turquie',
            r'216-71-125500',
            r'bt\s+net',
            r'achat\s+porteur\s+bt',
            r'reglement\s+achat\s+porteur\s+bt',
        ]
        
        # Mots-clés pour détecter les relevés
        self.releve_keywords = [
            r'relevé\s+de\s+compte',
            r'relevé\s+compte',
            r'statement\s+of\s+account',
            r'mouvement\s+de\s+compte',
        ]
        
        # Mots-clés pour détecter UBCI (extrait)
        self.ubci_keywords = [
            r'ubci',
            r'union\s+bancaire\s+pour\s+le\s+commerce',
            r'extrait\s+de\s+compte\s*',
            r'date\s+opération',
            r'natures\s+des\s+opérations',
        ]
        # Mots-clés pour détecter UIB (relevé)
        self.uib_keywords = [
            r'\buib\b',
            r'groupe\s+société\s+générale',
            r'relevé\s+de\s+compte',
            r'libellé\s+de\s+l\'opération',
            r'débit', r'crédit', r'valeur'
        ]
        
        # Compiler les expressions régulières
        self.zitouna_extrait_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.zitouna_extrait_keywords]
        self.amen_extrait_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.amen_extrait_keywords]
        self.releve_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.releve_keywords]
        self.btk_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.btk_keywords]
        self.bna_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.bna_keywords]
        self.bt_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.bt_keywords]
        self.ubci_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.ubci_keywords]
        self.uib_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.uib_keywords]

    def detect_document_type(self, pdf_path: str) -> Dict[str, any]:
        """
        Détecte le type de document PDF
        
        Returns:
            Dict avec les informations de détection:
            {
                'type': 'extrait_zitouna' | 'extrait_amen' | 'releve_zitouna' | 'releve_amen' | 'unknown',
                'bank': 'zitouna' | 'amen' | 'unknown',
                'document_type': 'extrait' | 'releve' | 'unknown',
                'confidence': float (0.0 à 1.0),
                'matched_keywords': list,
                'error': str (si erreur)
            }
        """
        try:
            # Extraire le texte du PDF
            text_content = self._extract_pdf_text(pdf_path)
            
            if not text_content:
                return {
                    'type': 'unknown',
                    'bank': 'unknown',
                    'document_type': 'unknown',
                    'confidence': 0.0,
                    'matched_keywords': [],
                    'error': 'Impossible d\'extraire le texte du PDF'
                }
            
            # Analyser le contenu
            detection_result = self._analyze_content(text_content)
            
            return detection_result
            
        except Exception as e:
            return {
                'type': 'unknown',
                'bank': 'unknown',
                'document_type': 'unknown',
                'confidence': 0.0,
                'matched_keywords': [],
                'error': f'Erreur lors de la détection: {str(e)}'
            }

    def _extract_pdf_text(self, pdf_path: str) -> str:
        """Extrait le texte du PDF"""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                text_parts = []
                # Analyser les premières pages (généralement suffisant pour la détection)
                for i, page in enumerate(pdf.pages[:3]):  # Limiter aux 3 premières pages
                    page_text = page.extract_text()
                    if page_text:
                        text_parts.append(page_text)
                
                return '\n'.join(text_parts)
        except Exception as e:
            print(f"Erreur lors de l'extraction du texte: {e}")
            return ""

    def _analyze_content(self, text: str) -> Dict[str, any]:
        """Analyse le contenu textuel pour déterminer le type de document"""
        
        # Normaliser le texte
        normalized_text = text.lower().replace('\n', ' ').replace('\r', ' ')
        
        # Scores de détection
        zitouna_score = 0
        amen_score = 0
        btk_score = 0
        bna_score = 0
        bt_score = 0
        ubci_score = 0
        uib_score = 0
        releve_score = 0
        
        matched_keywords = []
        
        # Détecter Zitouna
        for pattern in self.zitouna_extrait_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                zitouna_score += len(matches) * 2  # Poids plus élevé pour les extraits
                matched_keywords.extend(matches)
        
        # Détecter Amen Bank
        for pattern in self.amen_extrait_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                amen_score += len(matches) * 2  # Poids plus élevé pour les extraits
                matched_keywords.extend(matches)
        
        # Détecter les relevés
        for pattern in self.releve_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                releve_score += len(matches)
                matched_keywords.extend(matches)

        # Détecter BTK
        for pattern in self.btk_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                btk_score += len(matches) * 2
                matched_keywords.extend(matches)
        
        # Détecter BNA
        for pattern in self.bna_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                bna_score += len(matches) * 2
                matched_keywords.extend(matches)
        
        # Détecter BT
        for pattern in self.bt_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                bt_score += len(matches) * 2
                matched_keywords.extend(matches)
        
        # Détecter UBCI
        for pattern in self.ubci_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                ubci_score += len(matches) * 2
                matched_keywords.extend(matches)
        # Détecter UIB
        for pattern in self.uib_patterns:
            matches = pattern.findall(normalized_text)
            if matches:
                uib_score += len(matches) * 2
                matched_keywords.extend(matches)
        
        # Détection supplémentaire basée sur des patterns spécifiques
        # Pour Zitouna: chercher des patterns de montants avec point comme séparateur de milliers
        zitouna_amount_pattern = re.compile(r'\d+\.\d{3},\d{3}')
        zitouna_amounts = zitouna_amount_pattern.findall(text)
        if len(zitouna_amounts) > 2:  # Si on trouve plusieurs montants au format Zitouna
            zitouna_score += 1
        
        # Pour Amen Bank: chercher des patterns de montants avec espace comme séparateur de milliers
        amen_amount_pattern = re.compile(r'\d+\s\d{3},\d{3}')
        amen_amounts = amen_amount_pattern.findall(text)
        if len(amen_amounts) > 2:  # Si on trouve plusieurs montants au format Amen
            amen_score += 1
        
        # Déterminer le résultat
        max_score = max(zitouna_score, amen_score, btk_score, bna_score, bt_score, ubci_score, uib_score, releve_score)
        
        if max_score == 0:
            return {
                'type': 'unknown',
                'bank': 'unknown',
                'document_type': 'unknown',
                'confidence': 0.0,
                'matched_keywords': [],
                'error': 'Aucun pattern reconnu'
            }
        
        # Calculer la confiance
        total_possible_score = 10  # Score maximum théorique
        confidence = min(max_score / total_possible_score, 1.0)
        
        # Déterminer le type de document
        if bt_score > max(zitouna_score, amen_score, btk_score, bna_score, ubci_score, releve_score):
            bank = 'bt'
            if releve_score > 0:
                doc_type = 'releve'
                doc_type_full = 'releve_bt'
            else:
                doc_type = 'extrait'
                doc_type_full = 'extrait_bt'
        elif ubci_score > max(zitouna_score, amen_score, btk_score, bna_score, bt_score, releve_score):
            bank = 'ubci'
            doc_type = 'extrait'
            doc_type_full = 'extrait_ubci'
        elif uib_score > max(zitouna_score, amen_score, btk_score, bna_score, bt_score, ubci_score):
            bank = 'uib'
            doc_type = 'releve'
            doc_type_full = 'releve_uib'
        elif bna_score > zitouna_score and bna_score > amen_score and bna_score > btk_score and bna_score > releve_score:
            bank = 'bna'
            if releve_score > 0:
                doc_type = 'releve'
                doc_type_full = 'releve_bna'
            else:
                doc_type = 'extrait'
                doc_type_full = 'extrait_bna'
        elif btk_score > zitouna_score and btk_score > amen_score and btk_score > releve_score:
            bank = 'btk'
            if releve_score > 0:
                doc_type = 'releve'
                doc_type_full = 'releve_btk'
            else:
                doc_type = 'extrait'
                doc_type_full = 'extrait_btk'
        elif zitouna_score > amen_score and zitouna_score > releve_score:
            bank = 'zitouna'
            if releve_score > 0:
                doc_type = 'releve'
                doc_type_full = 'releve_zitouna'
            else:
                doc_type = 'extrait'
                doc_type_full = 'extrait_zitouna'
        elif amen_score > zitouna_score and amen_score > releve_score:
            bank = 'amen'
            if releve_score > 0:
                doc_type = 'releve'
                doc_type_full = 'releve_amen'
            else:
                doc_type = 'extrait'
                doc_type_full = 'extrait_amen'
        else:
            # Par défaut, considérer comme relevé si on a des scores de relevé
            if releve_score > 0:
                if zitouna_score > amen_score:
                    bank = 'zitouna'
                    doc_type = 'releve'
                    doc_type_full = 'releve_zitouna'
                else:
                    bank = 'amen'
                    doc_type = 'releve'
                    doc_type_full = 'releve_amen'
            elif btk_score > 0:
                bank = 'btk'
                doc_type = 'releve'
                doc_type_full = 'releve_btk'
            else:
                return {
                    'type': 'unknown',
                    'bank': 'unknown',
                    'document_type': 'unknown',
                    'confidence': confidence,
                    'matched_keywords': matched_keywords,
                    'error': 'Type de document ambigu'
                }
        
        return {
            'type': doc_type_full,
            'bank': bank,
            'document_type': doc_type,
            'confidence': confidence,
            'matched_keywords': matched_keywords,
            'error': None
        }

    def get_detection_summary(self, detection_result: Dict[str, any]) -> str:
        """Retourne un résumé lisible de la détection"""
        if detection_result['type'] == 'unknown':
            return f"❓ Type inconnu (Confiance: {detection_result['confidence']:.1%})"
        
        bank_names = {
            'zitouna': 'Zitouna Bank',
            'amen': 'Amen Bank',
            'btk': 'BTK Bank'
        }
        
        doc_types = {
            'extrait': 'Extrait',
            'releve': 'Relevé'
        }
        
        bank = bank_names.get(detection_result['bank'], 'Inconnue')
        doc_type = doc_types.get(detection_result['document_type'], 'Inconnu')
        confidence = detection_result['confidence']
        
        return f"✅ {bank} - {doc_type} (Confiance: {confidence:.1%})"

# Fonction utilitaire pour usage direct
def detect_pdf_type(pdf_path: str) -> Dict[str, any]:
    """Fonction utilitaire pour détecter le type d'un PDF"""
    detector = PDFBankDetector()
    return detector.detect_document_type(pdf_path)

if __name__ == "__main__":
    # Test du détecteur
    import sys
    if len(sys.argv) > 1:
        pdf_path = sys.argv[1]
        detector = PDFBankDetector()
        result = detector.detect_document_type(pdf_path)
        print("Résultat de la détection:")
        print(f"Type: {result['type']}")
        print(f"Banque: {result['bank']}")
        print(f"Type de document: {result['document_type']}")
        print(f"Confiance: {result['confidence']:.1%}")
        print(f"Mots-clés trouvés: {result['matched_keywords']}")
        if result['error']:
            print(f"Erreur: {result['error']}")
    else:
        print("Usage: python pdf_detector.py <chemin_vers_pdf>")

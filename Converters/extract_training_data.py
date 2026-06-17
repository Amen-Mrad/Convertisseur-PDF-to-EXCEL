import pdfplumber
import pandas as pd
import re

# Mets ici le chemin vers un PDF que tu veux utiliser pour collecter des données d'entraînement
PDF_PATH = r"C:\Users\amenm\OneDrive\Bureau\BANQUES\releveBIAT5.pdf"


results = []

with pdfplumber.open(PDF_PATH) as pdf:
    for page in pdf.pages:
        words = page.extract_words()  # OCR intégré de pdfplumber
        for w in words:
            text = w['text'].replace(" ", "")

            # Détecter si c'est un montant (nombre)
            if re.match(r'^[0-9\.,]+$', text):
                x_center = (w['x0'] + w['x1']) / 2

                results.append({
                    "amount": text,
                    "x_center": x_center,
                    "page": page.page_number
                })

# Convertir en tableau
df = pd.DataFrame(results)

# Sauvegarder en CSV pour annoter à la main
df.to_csv("training_raw.csv", index=False)

print("CSV créé : training_raw.csv")

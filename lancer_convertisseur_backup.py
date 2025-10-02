#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de lancement pour les convertisseurs Attijari Bank
Permet de choisir entre EXTRAT et RELEVE
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os

# Configuration des styles modernes
class ModernStyles:
    # Couleurs modernes
    PRIMARY = '#3498db'      # Bleu moderne
    SECONDARY = '#2ecc71'    # Vert moderne
    SUCCESS = '#27ae60'      # Vert succès
    WARNING = '#f39c12'      # Orange
    DANGER = '#e74c3c'       # Rouge
    INFO = '#17a2b8'         # Bleu info
    LIGHT = '#ecf0f1'        # Gris clair
    DARK = '#2c3e50'         # Gris foncé
    WHITE = '#ffffff'        # Blanc
    
    # Couleurs des banques
    ATTIJARI = '#F58E27'
    BIAT = '#1e3a8a'
    BTK = '#059669'
    QNB = '#dc2626'
    AMEN = '#7c3aed'
    ZITOUNA = '#16a34a'

class ConvertisseurLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur Bancaire - Sélection")
        self.root.geometry("500x600")
        self.root.configure(bg='#f5f5f5')
        
        # Mode sombre
        self.dark_mode = False
        
        # Configuration du style ttk
        self.setup_ttk_styles()
        
        # Etat de sélection
        self.selected_bank = None  # 'attijari' | 'biat' | 'amen' | 'zitouna' | 'bna' (QNB et BTK masqués)
        self.selected_type = tk.StringVar(value="")  # 'releve' | 'extrait'
        
        # Références UI
        self.bank_buttons = {}
        self.types_frame = None
        self.types_inner = None
        self.start_button = None
        
        # Créer le système de scroll
        self.setup_scrollable_ui()
    
    def setup_ttk_styles(self):
        """Configure les styles ttk modernes"""
        style = ttk.Style()
        
        # Style pour les boutons de banque
        style.configure('Bank.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       padding=(20, 15),
                       relief='flat',
                       borderwidth=0)
        
        # Style pour les boutons de type
        style.configure('Type.TRadiobutton',
                       font=('Segoe UI', 10),
                       padding=(10, 8))
        
        # Style pour le bouton de démarrage
        style.configure('Start.TButton',
                       font=('Segoe UI', 12, 'bold'),
                       padding=(30, 15),
                       relief='flat',
                       borderwidth=0)
        
        # Style pour les frames
        style.configure('Card.TFrame',
                       relief='solid',
                       borderwidth=1)
        
    def setup_scrollable_ui(self):
        """Configure l'interface avec scroll"""
        # Créer le canvas principal avec scrollbar
        self.canvas = tk.Canvas(self.root, bg='#f5f5f5', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f5f5f5', width=480)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        # Créer la fenêtre du canvas centrée
        self._canvas_window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Pack le canvas et la scrollbar
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Bind la molette de la souris
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        
        # Créer l'interface dans le frame scrollable
        self.create_ui()
        
        # Centrer le contenu après création
        self.root.after(100, self._center_content)
    
    def _on_mousewheel(self, event):
        """Gestion du scroll avec la molette"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _center_content(self):
        """Centre le contenu horizontalement"""
        try:
            canvas_width = self.canvas.winfo_width()
            if canvas_width > 1:  # S'assurer que le canvas a une largeur
                self.canvas.coords(self._canvas_window_id, canvas_width/2, 0)
        except Exception:
            pass
    
    
    def update_scroll(self):
        """Met à jour la zone de scroll"""
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def create_ui(self):
        # Carte blanche principale centrée - Plus petite
        main_card = tk.Frame(self.scrollable_frame, bg='white', relief='flat', borderwidth=0)
        main_card.pack(pady=30, padx=30, fill='both', expand=True)
        
        # Titre principal - Plus petit
        title_label = tk.Label(main_card, text="Convertisseur Bancaire", 
                              font=("Segoe UI", 18, "bold"), bg='white', fg='#2c3e50')
        title_label.pack(pady=(20, 15))
        
        # Étapes avec style moderne - Plus petites
        steps_frame = tk.Frame(main_card, bg='white')
        steps_frame.pack(pady=(0, 20))
        
        # Badge 1 - Inactif (gris) - Plus petit
        badge1 = tk.Frame(steps_frame, bg='#bdc3c7', relief='flat', borderwidth=0)
        badge1.pack(side='left', padx=3)
        tk.Label(badge1, text="1", font=("Segoe UI", 8, "bold"), 
                bg='#bdc3c7', fg='white', width=1, height=1).pack(padx=6, pady=3)
        tk.Label(steps_frame, text="Banque", font=("Segoe UI", 9), 
                bg='white', fg='#7f8c8d').pack(side='left', padx=(3, 15))
        
        # Badge 2 - Inactif (gris) - Plus petit
        badge2 = tk.Frame(steps_frame, bg='#bdc3c7', relief='flat', borderwidth=0)
        badge2.pack(side='left', padx=3)
        tk.Label(badge2, text="2", font=("Segoe UI", 8, "bold"), 
                bg='#bdc3c7', fg='white', width=1, height=1).pack(padx=6, pady=3)
        tk.Label(steps_frame, text="Type", font=("Segoe UI", 9), 
                bg='white', fg='#7f8c8d').pack(side='left', padx=(3, 15))
        
        # Badge 3 - Inactif (gris) - Plus petit
        badge3 = tk.Frame(steps_frame, bg='#bdc3c7', relief='flat', borderwidth=0)
        badge3.pack(side='left', padx=3)
        tk.Label(badge3, text="3", font=("Segoe UI", 8, "bold"), 
                bg='#bdc3c7', fg='white', width=1, height=1).pack(padx=6, pady=3)
        tk.Label(steps_frame, text="Conversion", font=("Segoe UI", 9), 
                bg='white', fg='#7f8c8d').pack(side='left', padx=(3, 0))
        
        # Titre de section dans la carte - Plus petit
        banks_title = tk.Label(main_card, text="Sélectionnez votre banque", 
                              font=("Segoe UI", 12, "bold"), bg='white', fg='#2c3e50')
        banks_title.pack(pady=(0, 15))
        
        # Grille des boutons de banques dans la carte
        banks_frame = tk.Frame(main_card, bg='white')
        banks_frame.pack(pady=(0, 20))
        
        # Configuration des boutons avec icônes et couleurs - Style exact de l'image
        bank_configs = [
            ('attijari', 'ATTIJARI BANK', '#e67e22', '#FFFFFF'),  # Orange
            ('biat', 'BIAT BANK', '#3498db', '#FFFFFF'),  # Bleu
            ('amen', 'AMEN BANK', '#16a085', '#FFFFFF'),  # Teal
            ('zitouna', 'ZITOUNA BANK', '#27ae60', '#FFFFFF'),  # Vert
            ('bna', 'BNA BANK', '#27ae60', '#FFFFFF')  # Vert
        ]
        
        for i, (bank_id, text, bg_color, fg_color) in enumerate(bank_configs):
            row = i // 3  # Grille 3 colonnes pour 5 banques
            col = i % 3
            
            # Bouton de banque avec style plus petit
            self.bank_buttons[bank_id] = tk.Button(
                banks_frame, text=f"🏦 {text}", width=15, height=2,
                bg=bg_color, fg=fg_color, font=("Segoe UI", 9, "bold"),
                activebackground=bg_color, activeforeground=fg_color,
                relief='flat', borderwidth=0, cursor='hand2',
                command=lambda b=bank_id: self.select_bank(b)
            )
            self.bank_buttons[bank_id].grid(row=row, column=col, padx=8, pady=8, sticky='nsew')
            
            # Effet hover
            self.bank_buttons[bank_id].bind('<Enter>', lambda e, btn=self.bank_buttons[bank_id]: self.on_hover_enter(btn))
            self.bank_buttons[bank_id].bind('<Leave>', lambda e, btn=self.bank_buttons[bank_id]: self.on_hover_leave(btn))
        
        # Configuration de la grille 3 colonnes
        for i in range(3):
            banks_frame.grid_columnconfigure(i, weight=1)
        
        # Zone des types (cachée au départ) - Dans la carte
        self.types_frame = tk.Frame(main_card, bg='white')
        self.types_frame.pack(pady=(0, 20))
        
        self.types_inner = tk.Frame(self.types_frame, bg='white')
        # caché tant qu'aucune banque n'est choisie
        self.types_inner.pack_forget()
        
        # Titre de section pour les types - Plus petit
        type_title = tk.Label(self.types_inner, text="Choisissez le type de document", 
                             font=("Segoe UI", 11, "bold"), bg='white', fg='#2c3e50')
        type_title.pack(pady=(8, 10))
        
        # Frame pour les boutons radio avec style moderne
        radio_frame = tk.Frame(self.types_inner, bg='white')
        radio_frame.pack()
        
        self.rb_releve = ttk.Radiobutton(radio_frame, text="Relevé de compte", value='releve', 
                                        variable=self.selected_type, command=self.on_type_change,
                                        style='Type.TRadiobutton')
        self.rb_extrait = ttk.Radiobutton(radio_frame, text="Extrait de compte", value='extrait', 
                                         variable=self.selected_type, command=self.on_type_change,
                                         style='Type.TRadiobutton')
        self.rb_releve.pack(side='left', padx=20)
        self.rb_extrait.pack(side='left', padx=20)
        
        # Bouton démarrer moderne - Plus petit
        start_frame = tk.Frame(main_card, bg='white')
        start_frame.pack(pady=(0, 15), fill='x')
        
        self.start_button = tk.Button(start_frame, text="▶️ DÉMARRER LA CONVERSION", state='disabled',
                                      bg='#27ae60', fg='white', 
                                      font=("Segoe UI", 10, "bold"),
                                      height=2, relief='flat', borderwidth=0,
                                      cursor='hand2', command=self.demarrer,
                                      disabledforeground='white')
        self.start_button.pack(fill='x', pady=5)
        
        # Informations avec style moderne - Plus petit
        info_label = tk.Label(main_card, text="Les fichiers Excel seront créés dans votre dossier Téléchargements", 
                             font=("Segoe UI", 8), bg='white', fg='#7f8c8d')
        info_label.pack(pady=(0, 20))
    
    def toggle_dark_mode(self):
        """Bascule entre le mode sombre et clair"""
        self.dark_mode = not self.dark_mode
        
        if self.dark_mode:
            # Mode sombre
            self.root.configure(bg='#2c3e50')
            self.scrollable_frame.configure(bg='#2c3e50')
            self.dark_mode_button.config(text="☀️", bg='#2c3e50', fg='#ecf0f1')
            # Mettre à jour tous les éléments pour le mode sombre
            self.update_dark_mode_elements()
        else:
            # Mode clair
            self.root.configure(bg=ModernStyles.LIGHT)
            self.scrollable_frame.configure(bg=ModernStyles.LIGHT)
            self.dark_mode_button.config(text="🌙", bg=ModernStyles.LIGHT, fg=ModernStyles.DARK)
            # Restaurer le mode clair
            self.update_light_mode_elements()
    
    def update_dark_mode_elements(self):
        """Met à jour tous les éléments pour le mode sombre"""
        # Cette fonction serait étendue pour mettre à jour tous les widgets
        pass
    
    def update_light_mode_elements(self):
        """Met à jour tous les éléments pour le mode clair"""
        # Cette fonction serait étendue pour restaurer le mode clair
        pass
    
    # --- Effets visuels ---
    def on_hover_enter(self, button):
        """Aucun effet hover - couleur constante"""
        # Pas de changement de couleur
        pass
    
    def on_hover_leave(self, button):
        """Aucun effet hover - couleur constante"""
        # Pas de changement de couleur
        pass

    # --- Interactions UI ---
    def select_bank(self, bank_key):
        # Mettre à jour la sélection visuelle
        self.selected_bank = bank_key
        for key, btn in self.bank_buttons.items():
            # Garder toujours les couleurs originales - pas de changement visuel
            if key == 'attijari':
                btn.configure(bg='#e67e22', fg='#FFFFFF', relief='flat', borderwidth=0)
            elif key == 'biat':
                btn.configure(bg='#3498db', fg='#FFFFFF', relief='flat', borderwidth=0)
            elif key == 'amen':
                btn.configure(bg='#16a085', fg='#FFFFFF', relief='flat', borderwidth=0)
            elif key == 'zitouna':
                btn.configure(bg='#27ae60', fg='#FFFFFF', relief='flat', borderwidth=0)
            elif key == 'bna':
                btn.configure(bg='#27ae60', fg='#FFFFFF', relief='flat', borderwidth=0)
            else:
                btn.configure(bg='#95a5a6', fg='#FFFFFF', relief='flat', borderwidth=0)
        
        # Ajuster les types disponibles selon la banque
        self.selected_type.set("")
        self.rb_releve.configure(state='normal')
        self.rb_extrait.configure(state='normal')
        if bank_key == 'biat':
            # BIAT: seulement Relevé pour l'instant
            self.rb_extrait.configure(state='disabled')
        elif bank_key in ('zitouna', 'attijari', 'bna'):
            # Ces banques: Relevé et Extrait disponibles
            self.rb_extrait.configure(state='normal')
        elif bank_key == 'amen':
            # Bloquer l'extrait pour AMEN BANK
            self.rb_extrait.configure(state='disabled')
            if self.selected_type.get() == 'extrait':
                self.selected_type.set('')
        
        # Afficher la zone des types avec une petite "animation" (clignotement doux)
        self.show_types_with_pulse()
        
        # Désactiver démarrer tant que le type n'est pas choisi
        self.update_start_enabled()
        
        # Mettre à jour le scroll
        self.update_scroll()
    
    def show_types_with_pulse(self):
        # Affiche et pulse la zone types
        self.types_inner.pack()
        original_bg = 'white'  # Garder le fond blanc
        pulse_bg = '#ecf5ff'
        
        def pulse(step):
            if step >= 6:
                # Finir proprement avec fond blanc
                for w in (self.types_inner,):
                    w.configure(bg='white')
                # Mettre à jour le scroll à la fin
                self.update_scroll()
                return
            bg = pulse_bg if step % 2 == 0 else original_bg
            self.types_inner.configure(bg=bg)
            for child in self.types_inner.winfo_children():
                # synchroniser le bg des labels uniquement
                if isinstance(child, tk.Label):
                    child.configure(bg=bg)
            self.root.after(120, lambda: pulse(step + 1))
        pulse(0)
    
    def on_type_change(self):
        self.update_start_enabled()
    
    def update_start_enabled(self):
        # Toujours activer le bouton pour permettre les messages d'erreur
        self.start_button.configure(state='normal')
    
    # --- Lancement ---
    def demarrer(self):
        bank = self.selected_bank
        typ = self.selected_type.get()
        
        # Vérifications spécifiques avec messages d'erreur personnalisés
        if not bank:
            messagebox.showwarning("Sélection incomplète", "Veuillez choisir une banque.")
            return
        
        if typ not in ('releve', 'extrait'):
            messagebox.showwarning("Sélection incomplète", "Sélectionnez le type de document.")
            return
        
        # Table de routage vers les scripts
        # Note: QNB et BTK masqués de l'interface mais fichiers conservés
        route = {
            ('attijari', 'extrait'): 'attijari_converter.py',
            ('attijari', 'releve'): 'attijari_releve_converter.py',
            ('biat', 'releve'): 'biat_releve_converter.py',
            ('biat', 'extrait'): 'biat_extrait_converter.py',
            ('amen', 'releve'): 'amen_releve_converter.py',
            ('amen', 'extrait'): 'amen_extrait_converter.py',
            ('zitouna', 'releve'): 'zitouna_releve_converter.py',
            ('zitouna', 'extrait'): 'zitouna_extrait_converter.py',
            ('bna', 'releve'): 'bna_releve_converter.py',
            ('bna', 'extrait'): 'bna_extrait_converter.py',
        }
        target = route.get((bank, typ))
        if not target:
            messagebox.showerror("Non pris en charge", "Ce type n'est pas disponible pour la banque choisie.")
            return
        
        try:
            if os.path.exists(target):
                subprocess.Popen([sys.executable, target])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", f"Fichier {target} non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur: {e}")
    
    def lancer_extrait(self):
        """Lance le convertisseur EXTRAT"""
        try:
            if os.path.exists("attijari_converter.py"):
                subprocess.Popen([sys.executable, "attijari_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier attijari_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur EXTRAT: {e}")
    
    def lancer_releve_attijari(self):
        """Lance le convertisseur RELEVE Attijari"""
        try:
            if os.path.exists("attijari_releve_converter.py"):
                subprocess.Popen([sys.executable, "attijari_releve_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier attijari_releve_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur RELEVE Attijari: {e}")
    
    def lancer_releve_biat(self):
        """Lance le convertisseur RELEVE BIAT"""
        try:
            if os.path.exists("biat_releve_converter.py"):
                subprocess.Popen([sys.executable, "biat_releve_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier biat_releve_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur RELEVE BIAT: {e}")
    
    def lancer_extrat_biat(self):
        """Lance le convertisseur EXTRAT BIAT"""
        try:
            if os.path.exists("biat_extrait_converter.py"):
                subprocess.Popen([sys.executable, "biat_extrait_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier biat_extrait_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur EXTRAT BIAT: {e}")

def main():
    root = tk.Tk()
    app = ConvertisseurLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()
        
        # Ouvrir le dialogue de sélection de fichiers
        files = filedialog.askopenfilenames(
            title="Sélectionner les fichiers PDF",
            filetypes=[("Fichiers PDF", "*.pdf"), ("Tous les fichiers", "*.*")]
        )
        
        if files:
            # Lancer le script de conversion avec les fichiers sélectionnés
            try:
                import subprocess
                import sys
                
                # Lancer le script avec les fichiers en argument
                subprocess.Popen([sys.executable, target_script] + list(files))
                self.root.destroy()
            except Exception as e:
                messagebox.showerror("Erreur", f"Impossible de lancer la conversion: {e}")
    
    def back_to_bank_selection(self):
        """Retourne à l'interface de sélection de banque"""
        # Cacher l'interface de sélection de fichiers
        self.file_selection_frame.pack_forget()
        
        # Réafficher l'interface de sélection de banque
        self.scrollable_frame.pack(fill='both', expand=True)
        
    def lancer_extrait(self):
        """Lance le convertisseur EXTRAT"""
        try:
            if os.path.exists("attijari_converter.py"):
                subprocess.Popen([sys.executable, "attijari_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier attijari_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur EXTRAT: {e}")
    
    def lancer_releve_attijari(self):
        """Lance le convertisseur RELEVE Attijari"""
        try:
            if os.path.exists("attijari_releve_converter.py"):
                subprocess.Popen([sys.executable, "attijari_releve_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier attijari_releve_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur RELEVE Attijari: {e}")
    
    def lancer_releve_biat(self):
        """Lance le convertisseur RELEVE BIAT"""
        try:
            if os.path.exists("biat_releve_converter.py"):
                subprocess.Popen([sys.executable, "biat_releve_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier biat_releve_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur RELEVE BIAT: {e}")
    
    def lancer_extrat_biat(self):
        """Lance le convertisseur EXTRAT BIAT"""
        try:
            if os.path.exists("biat_extrait_converter.py"):
                subprocess.Popen([sys.executable, "biat_extrait_converter.py"])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", "Fichier biat_extrait_converter.py non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur EXTRAT BIAT: {e}")
    
    
def main():
    root = tk.Tk()
    app = ConvertisseurLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()

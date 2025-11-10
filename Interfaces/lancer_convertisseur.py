
"""
Script de lancement pour les convertisseurs bancaires
Permet de choisir entre EXTRAT et RELEVE pour différentes banques
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os

# --- Project paths after folder reorganization ---
# Ensure we can access converter modules/scripts inside 'Converters'
CURRENT_DIR = os.path.dirname(__file__)
BASE_DIR = os.path.abspath(os.path.join(CURRENT_DIR, os.pardir))
CONVERTERS_DIR = os.path.join(BASE_DIR, 'Converters')

if os.path.isdir(CONVERTERS_DIR) and CONVERTERS_DIR not in sys.path:
    sys.path.insert(0, CONVERTERS_DIR)



class ModernStyles:
    PRIMARY = '#3498db'      
    SECONDARY = '#2ecc71'    
    SUCCESS = '#27ae60'      
    WARNING = '#f39c12'      
    DANGER = '#e74c3c'       
    INFO = '#17a2b8'         
    LIGHT = '#ecf0f1'        
    DARK = '#2c3e50'         
    WHITE = '#ffffff'        

    
    AMEN = '#7c3aed'
    ZITOUNA = '#16a34a'
    STB = '#1E3A8A'  
    BT = '#1e40af'
    WIFAK = '#0b5fa5'
    BH = '#1a365d'  # Couleur sombre pour BH Bank   

class ConvertisseurLauncher:
    def __init__(self, root):
        self.root = root
        self.root.title("Convertisseur Bancaire - Sélection")
        self.root.geometry("500x600")
        self.root.configure(bg='#f5f5f5')
        
      
        self.dark_mode = False
        
        
        self.setup_ttk_styles()
        
      
       
        self.selected_bank = None 
        self.selected_type = tk.StringVar(value="")  
        
        
        self.bank_buttons = {}
        self.types_frame = None
        self.types_inner = None
        self.start_button = None
        
        
        self.setup_scrollable_ui()
    
    def setup_ttk_styles(self):
        """Configure les styles ttk modernes"""
        style = ttk.Style()
        
        
        style.configure('Bank.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       padding=(20, 15),
                       relief='flat',
                       borderwidth=0)
        
        
        style.configure('Type.TRadiobutton',
                       font=('Segoe UI', 10),
                       padding=(10, 8))
        
        
        style.configure('Start.TButton',
                       font=('Segoe UI', 12, 'bold'),
                       padding=(30, 15),
                       relief='flat',
                       borderwidth=0)
        
        
        style.configure('Card.TFrame',
                       relief='solid',
                       borderwidth=1)
        
    def setup_scrollable_ui(self):
        """Configure l'interface avec scroll"""
        
        self.canvas = tk.Canvas(self.root, bg='#f5f5f5', highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg='#f5f5f5', width=480)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        
        self._canvas_window_id = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="n")
        
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
    
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
        
       
        self.create_ui()   
        self.root.after(100, self._center_content)
    
    def _on_mousewheel(self, event):
        """Gestion du scroll avec la molette"""
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _center_content(self):
        """Centre le contenu horizontalement"""
        try:
            canvas_width = self.canvas.winfo_width()
            if canvas_width > 1:  
                self.canvas.coords(self._canvas_window_id, canvas_width/2, 0)
        except Exception:
            pass
    
    
    def update_scroll(self):
        """Met à jour la zone de scroll"""
        self.scrollable_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def create_ui(self):
        
        main_card = tk.Frame(self.scrollable_frame, bg='white', relief='flat', borderwidth=0)
        main_card.pack(pady=30, padx=30, fill='both', expand=True)
        
       
        title_label = tk.Label(main_card, text="Convertisseur Bancaire", 
                              font=("Segoe UI", 18, "bold"), bg='white', fg='#2c3e50')
        title_label.pack(pady=(20, 15))
        
        
        steps_frame = tk.Frame(main_card, bg='white')
        steps_frame.pack(pady=(0, 20))
        
       
        badge1 = tk.Frame(steps_frame, bg='#bdc3c7', relief='flat', borderwidth=0)
        badge1.pack(side='left', padx=3)
        tk.Label(badge1, text="1", font=("Segoe UI", 8, "bold"), 
                bg='#bdc3c7', fg='white', width=1, height=1).pack(padx=6, pady=3)
        tk.Label(steps_frame, text="Banque", font=("Segoe UI", 9), 
                bg='white', fg='#7f8c8d').pack(side='left', padx=(3, 15))
        
        
        badge2 = tk.Frame(steps_frame, bg='#bdc3c7', relief='flat', borderwidth=0)
        badge2.pack(side='left', padx=3)
        tk.Label(badge2, text="2", font=("Segoe UI", 8, "bold"), 
                bg='#bdc3c7', fg='white', width=1, height=1).pack(padx=6, pady=3)
        tk.Label(steps_frame, text="Type", font=("Segoe UI", 9), 
                bg='white', fg='#7f8c8d').pack(side='left', padx=(3, 15))
        
        
        badge3 = tk.Frame(steps_frame, bg='#bdc3c7', relief='flat', borderwidth=0)
        badge3.pack(side='left', padx=3)
        tk.Label(badge3, text="3", font=("Segoe UI", 8, "bold"), 
                bg='#bdc3c7', fg='white', width=1, height=1).pack(padx=6, pady=3)
        tk.Label(steps_frame, text="Conversion", font=("Segoe UI", 9), 
                bg='white', fg='#7f8c8d').pack(side='left', padx=(3, 0))
        
        
        banks_title = tk.Label(main_card, text="Sélectionnez votre banque", 
                              font=("Segoe UI", 12, "bold"), bg='white', fg='#2c3e50')
        banks_title.pack(pady=(0, 15))
        
        
        banks_frame = tk.Frame(main_card, bg='white')
        banks_frame.pack(pady=(0, 20))
        
        
        bank_configs = [
            ('amen', 'AMEN BANK', '#16a085', '#FFFFFF'), 
            ('zitouna', 'ZITOUNA BANK', '#27ae60', '#FFFFFF'), 
            ('bna', 'BNA BANK', '#27ae60', '#FFFFFF'),  
            ('stb', 'STB BANK', '#0082ca', '#000000'),  # Fond bleu, texte noir
            ('wifak', 'WIFAK BANK', '#8b939a', '#0658a0'),  # Fond gris, texte bleu
            ('biat', 'BIAT BANK', '#004578', '#f28800'),  # Fond bleu foncé, texte orange
            ('bt', 'BANQUE DE TUNISIE', '#1e40af', '#FFFFFF'), 
            ('ubci', 'UBCI BANK', '#d8f0f1', '#0082ca'),  # Fond bleu, texte bleu
            # ('qnb', 'QNB BANK', '#1a365d', '#FFFFFF'),  # QNB temporairement caché
        ]
        
        for i, (bank_id, text, bg_color, fg_color) in enumerate(bank_configs):
            row = i // 3  
            col = i % 3
                      
            self.bank_buttons[bank_id] = tk.Button(
                banks_frame, text=f"🏦 {text}", width=15, height=2,
                bg=bg_color, fg=fg_color, font=("Segoe UI", 9, "bold"),
                activebackground=bg_color, activeforeground=fg_color,
                relief='solid', borderwidth=1, highlightthickness=1,
                highlightcolor='black', highlightbackground='black',
                cursor='hand2', command=lambda b=bank_id: self.select_bank(b)
            )
            self.bank_buttons[bank_id].grid(row=row, column=col, padx=8, pady=8, sticky='nsew')
            
           
            self.bank_buttons[bank_id].bind('<Enter>', lambda e, btn=self.bank_buttons[bank_id]: self.on_hover_enter(btn))
            self.bank_buttons[bank_id].bind('<Leave>', lambda e, btn=self.bank_buttons[bank_id]: self.on_hover_leave(btn))
        
       
        for i in range(3):
            banks_frame.grid_columnconfigure(i, weight=1)
        
       
        self.types_frame = tk.Frame(main_card, bg='white')
        self.types_frame.pack(pady=(0, 20))
        
        self.types_inner = tk.Frame(self.types_frame, bg='white')
        
        self.types_inner.pack_forget()
        
       
        type_title = tk.Label(self.types_inner, text="Choisissez le type de document", 
                             font=("Segoe UI", 11, "bold"), bg='white', fg='#2c3e50')
        type_title.pack(pady=(8, 10))
        
        
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
        
       
        start_frame = tk.Frame(main_card, bg='white')
        start_frame.pack(pady=(0, 15), fill='x')
        
        self.start_button = tk.Button(start_frame, text="▶️ DÉMARRER LA CONVERSION", state='disabled',
                                      bg='#27ae60', fg='white', 
                                      font=("Segoe UI", 10, "bold"),
                                      height=2, relief='flat', borderwidth=0,
                                      cursor='hand2', command=self.demarrer,
                                      disabledforeground='white')
        self.start_button.pack(fill='x', pady=5)
        
        
        info_label = tk.Label(main_card, text="Les fichiers Excel seront créés dans votre dossier Téléchargements", 
                             font=("Segoe UI", 8), bg='white', fg='#7f8c8d')
        info_label.pack(pady=(0, 20))
    
    def toggle_dark_mode(self):
        """Bascule entre le mode sombre et clair"""
        self.dark_mode = not self.dark_mode
        
        if self.dark_mode:
           
            self.root.configure(bg='#2c3e50')
            self.scrollable_frame.configure(bg='#2c3e50')
            self.dark_mode_button.config(text="☀️", bg='#2c3e50', fg='#ecf0f1')
           
            self.update_dark_mode_elements()
        else:
           
            self.root.configure(bg=ModernStyles.LIGHT)
            self.scrollable_frame.configure(bg=ModernStyles.LIGHT)
            self.dark_mode_button.config(text="🌙", bg=ModernStyles.LIGHT, fg=ModernStyles.DARK)
           
            self.update_light_mode_elements()
    
    def update_dark_mode_elements(self):
        """Met à jour tous les éléments pour le mode sombre"""
       
        pass
    
    def update_light_mode_elements(self):
        """Met à jour tous les éléments pour le mode clair"""
        
        pass
    

    def on_hover_enter(self, button):
        """Aucun effet hover - couleur constante"""
      
        pass
    
    def on_hover_leave(self, button):
        """Aucun effet hover - couleur constante"""
      
        pass


    
    def select_bank(self, bank_key):

        self.selected_bank = bank_key
    
        for key, btn in self.bank_buttons.items():

            if key == 'biat':
                btn.configure(bg='#004578', fg='#f28800', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'amen':
                btn.configure(bg='#16a085', fg='#FFFFFF', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'zitouna':
                btn.configure(bg='#27ae60', fg='#FFFFFF', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'bna':
                btn.configure(bg='#27ae60', fg='#FFFFFF', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'stb':
                btn.configure(bg='#0082ca', fg='#000000', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'bt':
                btn.configure(bg='#1e40af', fg='#FFFFFF', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'wifak':
                btn.configure(bg='#8b939a', fg='#0658a0', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            elif key == 'ubci':
                btn.configure(bg='#d8f0f1', fg='#0082ca', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            # elif key == 'qnb':
            #     btn.configure(bg='#1a365d', fg='#FFFFFF', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
            else:
                btn.configure(bg='#95a5a6', fg='#FFFFFF', relief='solid', borderwidth=1, highlightthickness=1, highlightcolor='black', highlightbackground='black')
        

        self.selected_type.set("")
        self.rb_releve.configure(state='normal')
        self.rb_extrait.configure(state='normal')
       
        if bank_key == 'biat':

            self.rb_extrait.configure(state='disabled')
        elif bank_key in ('zitouna', 'bna'):

            self.rb_extrait.configure(state='normal')
        elif bank_key == 'amen':
            # AMEN Bank - supporte les deux types
            self.rb_extrait.configure(state='normal')
            self.rb_releve.configure(state='normal')
        elif bank_key == 'ubci':
            # UBCI Bank - seulement extrait
            self.rb_releve.configure(state='disabled')
            self.rb_extrait.configure(state='normal')
            if self.selected_type.get() == 'releve':
                self.selected_type.set('')
        elif bank_key == 'stb':
            
            self.rb_extrait.configure(state='normal')
        elif bank_key == 'bt':

            self.rb_extrait.configure(state='disabled')
            if self.selected_type.get() == 'extrait':
                self.selected_type.set('')
      
        elif bank_key == 'wifak':
            self.rb_extrait.configure(state='normal')
        elif bank_key == 'ubci':
            # UBCI Bank - supporte les deux types
            self.rb_extrait.configure(state='normal')
            self.rb_releve.configure(state='normal')
        # elif bank_key == 'qnb':
        #     # QNB Bank - seulement relevé
        #     self.rb_releve.configure(state='normal')
        #     self.rb_extrait.configure(state='disabled')
        #     if self.selected_type.get() == 'extrait':
        #         self.selected_type.set('')
        else:

            try:
                if not self.rb_releve.winfo_manager():
                    self.rb_releve.pack(side='left', padx=20)
            except Exception:
                pass
        

        self.show_types_with_pulse()
        

        self.update_start_enabled()
        

        self.update_scroll()
    
    def show_types_with_pulse(self):

        self.types_inner.pack()
        original_bg = 'white' 
        pulse_bg = '#ecf5ff'
        
        def pulse(step):
          
            if step >= 6:

                for w in (self.types_inner,):

                    w.configure(bg='white')

                self.update_scroll()

                return
            bg = pulse_bg if step % 2 == 0 else original_bg
            self.types_inner.configure(bg=bg)
            for child in self.types_inner.winfo_children():

                if isinstance(child, tk.Label):

                    child.configure(bg=bg)
            self.root.after(120, lambda: pulse(step + 1))
        pulse(0)
    
    def on_type_change(self):
        self.update_start_enabled()
    
    

   
    def update_start_enabled(self):
        
        self.start_button.configure(state='normal')
    

    def demarrer(self):
        bank = self.selected_bank
        typ = self.selected_type.get()
        

        if not bank:
            messagebox.showwarning("Sélection incomplète", "Veuillez choisir une banque.")
            return
        
        if typ not in ('releve', 'extrait'):
            messagebox.showwarning("Sélection incomplète", "Sélectionnez le type de document.")
            return
        
        
        route = {
            ('biat', 'releve'): 'biat_releve_converter.py',
            ('biat', 'extrait'): 'biat_extrait_converter.py',
            ('amen', 'releve'): 'amen_releve_converter.py',
            ('amen', 'extrait'): 'amen_extrait_converter.py',
            ('zitouna', 'releve'): 'zitouna_releve_converter.py',
            ('zitouna', 'extrait'): 'zitouna_extrait_converter.py',
            ('bna', 'releve'): 'bna_releve_converter.py',
            ('bna', 'extrait'): 'bna_extrait_converter.py',
            ('stb', 'releve'): 'stb_releve_converter.py',
            ('stb', 'extrait'): 'stb_extrait_converter.py',
            ('bt', 'releve'): 'bt_releve_converter.py',
            ('wifak', 'extrait'): 'wifak_extrait_converter.py',
            ('wifak', 'releve'): 'wifak_releve_converter.py',
            ('ubci', 'releve'): 'ubci_releve_converter.py',
            ('ubci', 'extrait'): 'ubci_extrait_converter.py',
            # ('qnb', 'releve'): 'qnb_releve_converter.py',  # QNB temporairement caché
      
        }
        target = route.get((bank, typ))
        if not target:
            messagebox.showerror("Non pris en charge", f"Ce type ({typ}) n'est pas disponible pour la banque ({bank}) choisie.")
            return
        
        try:
            target_path = os.path.join(CONVERTERS_DIR, target)
            if os.path.exists(target_path):
                subprocess.Popen([sys.executable, target_path])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", f"Fichier {target_path} non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur: {e}")
    
    
    def lancer_releve_biat(self):
        """Lance le convertisseur RELEVE BIAT"""
        try:
            target = os.path.join(CONVERTERS_DIR, "biat_releve_converter.py")
            if os.path.exists(target):
                subprocess.Popen([sys.executable, target])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", f"Fichier {target} non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur RELEVE BIAT: {e}")
    
    def lancer_extrat_biat(self):
        """Lance le convertisseur EXTRAT BIAT"""
        try:
            target = os.path.join(CONVERTERS_DIR, "biat_extrait_converter.py")
            if os.path.exists(target):
                subprocess.Popen([sys.executable, target])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", f"Fichier {target} non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur EXTRAT BIAT: {e}")
    
    def lancer_releve_stb(self):
        """Lance le convertisseur RELEVE STB"""
        try:
            target = os.path.join(CONVERTERS_DIR, "stb_releve_converter.py")
            if os.path.exists(target):
                subprocess.Popen([sys.executable, target])
                self.root.destroy()
            else:
                messagebox.showerror("Erreur", f"Fichier {target} non trouvé")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lancer le convertisseur RELEVE STB: {e}")

def main():
    root = tk.Tk()
    app = ConvertisseurLauncher(root)
    root.mainloop()

if __name__ == "__main__":
    main()

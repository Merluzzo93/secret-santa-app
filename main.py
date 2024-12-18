import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import random
import time
from PIL import Image, ImageTk
import os
import openpyxl
from typing import List, Dict

class SecretSantaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Annaelle Bags Secret Santa")
        
        # Dimensioni finestra
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        window_width = int(screen_width * 0.8)
        window_height = int(screen_height * 0.8)
        
        # Centra la finestra
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.configure(bg="#EDE4DA")
        
        # Colori
        self.colors = {
            'brand': '#D17272',      # Annaelle Bags
            'secondary': '#6A8D6B',  # Secret Santa
            'background': '#EDE4DA', # Sfondo
            'text': '#333333'        # Testo generale
        }
        
        # Lista partecipanti e vincitori
        self.partecipanti: List[Dict[str, str]] = []
        self.vincitori: List[Dict[str, str]] = []
        self.current_draw = 1  # Tiene traccia dell'estrazione corrente (1=terzo, 2=secondo, 3=primo)
        
        self.setup_styles()
        self.create_main_interface()
        
        # Configura il ridimensionamento
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
    def setup_styles(self):
        """Configura gli stili dell'applicazione"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configura lo stile generale
        style.configure(
            'Main.TFrame',
            background=self.colors['background']
        )
        
        # Configura lo stile dei pulsanti
        style.configure(
            'Draw.TButton',
            font=('League Spartan', 12, 'bold'),
            padding=10,
            background=self.colors['secondary']
        )
        
        # Configura lo stile delle etichette
        style.configure(
            'Title.TLabel',
            font=('League Spartan', 28, 'bold'),
            foreground=self.colors['brand'],
            background=self.colors['background']
        )
        
        style.configure(
            'Subtitle.TLabel',
            font=('League Spartan', 16),
            foreground=self.colors['secondary'],
            background=self.colors['background']
        )

    def create_main_interface(self):
        """Crea l'interfaccia principale"""
        main_frame = ttk.Frame(self.root, style='Main.TFrame')
        main_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=20)
        main_frame.grid_columnconfigure(0, weight=1)
        
        # Titolo
        title_label = ttk.Label(
            main_frame,
            text="ANNAELLE BAGS",
            style='Title.TLabel'
        )
        title_label.grid(row=0, column=0, pady=(0, 5))
        
        subtitle_label = ttk.Label(
            main_frame,
            text="SECRET SANTA",
            style='Subtitle.TLabel'
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))
        
        # Pulsante per caricare Excel
        upload_button = ttk.Button(
            main_frame,
            text="Carica Lista Partecipanti (XLSX)",
            command=self.carica_excel,
            style='Draw.TButton'
        )
        upload_button.grid(row=2, column=0, pady=20)
        
        # Lista partecipanti
        list_frame = ttk.Frame(main_frame)
        list_frame.grid(row=3, column=0, sticky='nsew', pady=20)
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        self.partecipanti_listbox = tk.Listbox(
            list_frame,
            font=('League Spartan', 12),
            bg=self.colors['background'],
            fg=self.colors['text'],
            selectmode='none',
            height=15
        )
        self.partecipanti_listbox.grid(row=0, column=0, sticky='nsew')
        
        scrollbar = ttk.Scrollbar(list_frame, orient='vertical', command=self.partecipanti_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')
        self.partecipanti_listbox.configure(yscrollcommand=scrollbar.set)
        
        # Pulsante Estrazione
        self.extract_button = ttk.Button(
            main_frame,
            text="Estrai Terzo Posto",
            command=self.avvia_estrazione,
            style='Draw.TButton'
        )
        self.extract_button.grid(row=4, column=0, pady=20)
        
    def carica_excel(self):
        """Carica la lista partecipanti da file Excel"""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if filename:
            try:
                wb = openpyxl.load_workbook(filename)
                ws = wb.active
                
                self.partecipanti = []
                self.partecipanti_listbox.delete(0, tk.END)
                
                for row in ws.iter_rows(min_row=2):  # Salta l'intestazione
                    if row[0].value and row[1].value:  # Controlla che ci siano nome e social
                        nome = str(row[0].value).strip()
                        social = str(row[1].value).strip()
                        self.partecipanti.append({'nome': nome, 'social': social})
                        self.partecipanti_listbox.insert(tk.END, f"{nome} - {social}")
                
                messagebox.showinfo("Successo", "Lista partecipanti caricata correttamente!")
                
            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel caricamento del file: {str(e)}")
    
    def avvia_estrazione(self):
        """Gestisce l'estrazione dei vincitori"""
        if not self.partecipanti:
            messagebox.showerror("Errore", "Nessun partecipante nella lista!")
            return
            
        # Rimuove i vincitori precedenti dalla lista dei possibili estratti
        disponibili = [p for p in self.partecipanti if p not in self.vincitori]
        
        if not disponibili:
            messagebox.showerror("Errore", "Non ci sono più partecipanti disponibili!")
            return
            
        # Crea finestra per l'animazione
        win = tk.Toplevel(self.root)
        win.title("Estrazione in corso...")
        
        # Dimensiona e posiziona la finestra
        width = 800
        height = 600
        x = (self.root.winfo_screenwidth() - width) // 2
        y = (self.root.winfo_screenheight() - height) // 2
        win.geometry(f"{width}x{height}+{x}+{y}")
        
        # Canvas per l'animazione
        canvas = tk.Canvas(win, width=width, height=height, bg=self.colors['background'])
        canvas.pack(fill=tk.BOTH, expand=True)
        
        # Seleziona il vincitore
        vincitore = random.choice(disponibili)
        self.vincitori.append(vincitore)
        
        # Determina quale immagine usare
        img_paths = {
            1: "terzo-posto.png",
            2: "secondo-posto.png",
            3: "primo-posto.png"
        }
        
        # Carica e mostra l'immagine del posto
        try:
            img = Image.open(img_paths[self.current_draw])
            img = img.resize((width, height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(img)
            canvas.create_image(0, 0, anchor='nw', image=photo)
            win.photo = photo  # Mantiene un riferimento
        except Exception as e:
            messagebox.showerror("Errore", f"Impossibile caricare l'immagine: {str(e)}")
            win.destroy()
            return
            
        # Animazione nomi
        def animate_names(count=0):
            if count < 20:  # Numero di cambi prima di mostrare il vincitore
                canvas.delete('nome')
                nome = random.choice(disponibili)['nome']
                canvas.create_text(
                    width/2,
                    height-30,
                    text=nome,
                    font=('League Spartan', 24, 'bold'),
                    fill='white',
                    tags='nome'
                )
                win.after(100, lambda: animate_names(count + 1))
            else:
                # Mostra il vincitore
                canvas.delete('nome')
                canvas.create_text(
                    width/2,
                    height-30,
                    text=f"{vincitore['nome']} - {vincitore['social']}",
                    font=('League Spartan', 24, 'bold'),
                    fill='white',
                    tags='nome'
                )
                
                # Aggiorna il pulsante per la prossima estrazione
                self.current_draw += 1
                if self.current_draw <= 3:
                    next_text = {
                        2: "Estrai Secondo Posto",
                        3: "Estrai Primo Posto"
                    }
                    self.extract_button.configure(text=next_text[self.current_draw])
                else:
                    self.extract_button.configure(state='disabled')
                
        animate_names()

if __name__ == "__main__":
    root = tk.Tk()
    app = SecretSantaApp(root)
    root.mainloop()
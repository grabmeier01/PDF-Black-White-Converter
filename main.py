import os
import sys
import fitz  # PyMuPDF
from PIL import Image
import io
import img2pdf
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import webbrowser
import platform
import datetime
import json
import logging
from logging.handlers import RotatingFileHandler

# Logger konfigurieren
log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
log_file = "pdf_converter.log"

log_handler = RotatingFileHandler(log_file, maxBytes=5*1024*1024, backupCount=2)
log_handler.setFormatter(log_formatter)

logger = logging.getLogger()
logger.setLevel(logging.INFO)
logger.addHandler(log_handler)

# Globale Einstellungen
SETTINGS_FILE = "pdf_converter_settings.json"

# Standardeinstellungen
DEFAULT_SETTINGS = {
    "threshold": 180,
    "dpi": 300,
    "open_after": True,
    "output_dir": "",
    "output_suffix": "_SW",
    "mode": "bw",  # 'bw' oder 'grayscale'
    "page_range": "",
    "overwrite": "ask",  # 'ask', 'overwrite', 'skip'
    "compression": 95
}

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, 'r') as f:
                settings = json.load(f)
                # Stelle sicher, dass alle Einstellungen vorhanden sind
                for key, value in DEFAULT_SETTINGS.items():
                    if key not in settings:
                        settings[key] = value
                return settings
        except:
            return DEFAULT_SETTINGS.copy()
    else:
        return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    with open(SETTINGS_FILE, 'w') as f:
        json.dump(settings, f, indent=4)

# Hilfsfunktion zur Analyse des Seitenbereichs
def parse_page_range(page_range_str, total_pages):
    if not page_range_str.strip():
        return list(range(total_pages))
    
    pages = []
    parts = page_range_str.split(',')
    for part in parts:
        part = part.strip()
        if '-' in part:
            start_end = part.split('-')
            if len(start_end) == 2:
                start, end = start_end
                start = int(start.strip()) if start.strip() else 1
                end = int(end.strip()) if end.strip() else total_pages
                start = max(1, start)
                end = min(total_pages, end)
                pages.extend(range(start-1, end))
        else:
            if part.strip():
                try:
                    page = int(part.strip())
                    if 1 <= page <= total_pages:
                        pages.append(page-1)
                except ValueError:
                    pass
    # Duplikate entfernen und sortieren
    return sorted(set(pages))

def convert_pdf_to_bw(input_pdf, output_pdf, threshold=150, dpi=300, progress_callback=None, page_range=None, mode='bw', compression=95):
    """
    Konvertiert eine PDF in eine Schwarz-Weiß- oder Graustufen-PDF
    :param input_pdf: Pfad zur Eingabe-PDF
    :param output_pdf: Pfad für Ausgabe-PDF
    :param threshold: Schwellenwert (0-255), niedrig = dunkler (nur für BW)
    :param dpi: Auflösung für Bildkonvertierung
    :param progress_callback: Callback für Fortschrittsupdates
    :param page_range: Liste der zu verarbeitenden Seiten (0-basiert)
    :param mode: 'bw' für Schwarz-Weiß, 'grayscale' für Graustufen
    :param compression: Kompressionsstufe (1-100)
    """
    if progress_callback:
        progress_callback(0, f"Überprüfe Datei: {os.path.basename(input_pdf)}")
    
    if not os.path.exists(input_pdf):
        error_msg = f"FEHLER: Datei nicht gefunden: {input_pdf}"
        if progress_callback:
            progress_callback(0, error_msg)
        logger.error(error_msg)
        return False
    
    start_time = time.time()
    success = False
    bw_images = []
    
    try:
        doc = fitz.open(input_pdf)
        total_pages = len(doc)
        
        # Bestimme die zu verarbeitenden Seiten
        if page_range is None:
            page_range = list(range(total_pages))
        else:
            # Filtere ungültige Seiten
            page_range = [p for p in page_range if p < total_pages]
        
        if not page_range:
            error_msg = "Keine gültigen Seiten zum Konvertieren gefunden."
            if progress_callback:
                progress_callback(0, error_msg)
            logger.error(error_msg)
            return False
        
        num_pages = len(page_range)
        
        for idx, page_num in enumerate(page_range):
            if progress_callback:
                status = f"Verarbeite Seite {idx+1}/{num_pages} (Seite {page_num+1})"
                progress_callback((idx+1)/num_pages * 100, status)
            
            page = doc.load_page(page_num)
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            img_data = pix.tobytes("ppm")
            
            with Image.open(io.BytesIO(img_data)) as img:
                if mode == 'bw':
                    gray_img = img.convert('L')
                    bw_img = gray_img.point(lambda x: 0 if x < threshold else 255, '1')
                    img_bytes = io.BytesIO()
                    bw_img.save(img_bytes, format='TIFF', compression='group4')
                else:  # Graustufen
                    gray_img = img.convert('L')
                    img_bytes = io.BytesIO()
                    gray_img.save(img_bytes, format='JPEG', quality=compression)
                
                bw_images.append(img_bytes.getvalue())
        
        # Erstelle PDF aus den Bildern
        with open(output_pdf, "wb") as f:
            f.write(img2pdf.convert(bw_images))
        
        duration = time.time() - start_time
        success_msg = (
            f"✅ Konvertierung erfolgreich abgeschlossen in {duration:.1f} Sekunden!\n"
            f"Eingabe: {os.path.basename(input_pdf)}\n"
            f"Ausgabe: {os.path.basename(output_pdf)}\n"
            f"Dateigröße: {os.path.getsize(output_pdf)/1024:.1f} KB"
        )
        
        if progress_callback:
            progress_callback(100, success_msg)
        logger.info(success_msg)
        success = True
        
    except Exception as e:
        error_msg = f"FEHLER während der Konvertierung: {str(e)}"
        if progress_callback:
            progress_callback(0, error_msg)
        logger.error(error_msg, exc_info=True)
    
    return success

class PDFConverterApp:
    def __init__(self, root):
        self.root = root
        self.settings = load_settings()
        self.root.title("PDF zu Schwarz-Weiß Konverter")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        # Variablen initialisieren
        self.input_files = []
        self.current_file_index = 0
        self.total_files = 0
        
        self.threshold_value = tk.IntVar(value=self.settings['threshold'])
        self.dpi_value = tk.StringVar(value=str(self.settings['dpi']))
        self.open_after_var = tk.BooleanVar(value=self.settings['open_after'])
        self.page_range_var = tk.StringVar(value=self.settings['page_range'])
        self.overwrite_var = tk.StringVar(value=self.settings['overwrite'])
        self.mode_var = tk.StringVar(value=self.settings['mode'])
        self.compression_var = tk.IntVar(value=self.settings['compression'])
        self.timestamp_var = tk.BooleanVar(value=False)
        
        # Stilkonfiguration
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6)
        self.style.configure("TLabel", padding=6)
        self.style.configure("TFrame", padding=10)
        
        self.create_widgets()
    
    def create_widgets(self):
        # Haupt-Notebook für Tabs
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Haupt-Tab
        main_frame = ttk.Frame(notebook)
        notebook.add(main_frame, text="Konvertierung")
        
        # Log-Tab
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="Protokoll")
        
        # Haupt-Tab Inhalt
        self.create_main_tab(main_frame)
        
        # Log-Tab Inhalt
        self.create_log_tab(log_frame)
    
    def create_main_tab(self, parent):
        # Eingabedateien
        input_frame = ttk.LabelFrame(parent, text="Eingabe-PDFs")
        input_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Liste für PDF-Dateien
        self.input_listbox = tk.Listbox(input_frame, height=5)
        self.input_listbox.pack(fill=tk.X, padx=5, pady=5, expand=True)
        
        scrollbar = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=self.input_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.input_listbox.config(yscrollcommand=scrollbar.set)
        
        btn_frame = ttk.Frame(input_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(btn_frame, text="Dateien hinzufügen", command=self.browse_input).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Alle entfernen", command=self.clear_input).pack(side=tk.LEFT, padx=5)
        
        # Ausgabeeinstellungen
        output_frame = ttk.LabelFrame(parent, text="Ausgabeeinstellungen")
        output_frame.pack(fill=tk.X, padx=10, pady=5)
        
        output_dir_frame = ttk.Frame(output_frame)
        output_dir_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_dir_frame, text="Ausgabeordner:").pack(side=tk.LEFT)
        self.output_dir_entry = ttk.Entry(output_dir_frame)
        self.output_dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.output_dir_entry.insert(0, self.settings['output_dir'])
        ttk.Button(output_dir_frame, text="Durchsuchen", command=self.browse_output_dir).pack(side=tk.RIGHT)
        
        suffix_frame = ttk.Frame(output_frame)
        suffix_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(suffix_frame, text="Dateinamensuffix:").pack(side=tk.LEFT)
        self.suffix_entry = ttk.Entry(suffix_frame, width=15)
        self.suffix_entry.pack(side=tk.LEFT, padx=5)
        self.suffix_entry.insert(0, self.settings['output_suffix'])
        
        ttk.Checkbutton(
            suffix_frame, 
            text="Zeitstempel hinzufügen",
            variable=self.timestamp_var
        ).pack(side=tk.RIGHT, padx=10)
        
        # Konvertierungseinstellungen
        conv_frame = ttk.LabelFrame(parent, text="Konvertierungseinstellungen")
        conv_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Seitenbereich
        range_frame = ttk.Frame(conv_frame)
        range_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(range_frame, text="Seitenbereich (z.B. 1-5,8,10-12):").pack(side=tk.LEFT)
        self.page_range_entry = ttk.Entry(range_frame, width=25)
        self.page_range_entry.pack(side=tk.LEFT, padx=5)
        self.page_range_entry.insert(0, self.settings['page_range'])
        
        # Modus
        mode_frame = ttk.Frame(conv_frame)
        mode_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(mode_frame, text="Modus:").pack(side=tk.LEFT)
        ttk.Radiobutton(
            mode_frame, 
            text="Schwarz-Weiß", 
            variable=self.mode_var, 
            value="bw"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Radiobutton(
            mode_frame, 
            text="Graustufen", 
            variable=self.mode_var, 
            value="grayscale"
        ).pack(side=tk.LEFT, padx=5)
        
        # Kompression
        comp_frame = ttk.Frame(conv_frame)
        comp_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(comp_frame, text="Kompressionsstufe (1-100):").pack(side=tk.LEFT)
        comp_scale = ttk.Scale(
            comp_frame, 
            from_=1, 
            to=100, 
            variable=self.compression_var,
            command=self.update_compression_label
        )
        comp_scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.compression_label = ttk.Label(comp_frame, text=str(self.compression_var.get()), width=5)
        self.compression_label.pack(side=tk.RIGHT)
        
        # Überschreiben
        overwrite_frame = ttk.Frame(conv_frame)
        overwrite_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(overwrite_frame, text="Bei vorhandener Ausgabedatei:").pack(side=tk.LEFT)
        
        ttk.Radiobutton(
            overwrite_frame, 
            text="Fragen", 
            variable=self.overwrite_var, 
            value="ask"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Radiobutton(
            overwrite_frame, 
            text="Überschreiben", 
            variable=self.overwrite_var, 
            value="overwrite"
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Radiobutton(
            overwrite_frame, 
            text="Überspringen", 
            variable=self.overwrite_var, 
            value="skip"
        ).pack(side=tk.LEFT, padx=5)
        
        # Fortschrittsbalken
        self.progress_frame = ttk.Frame(parent)
        self.progress_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.progress_bar = ttk.Progressbar(
            self.progress_frame, 
            orient=tk.HORIZONTAL, 
            mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, expand=True)
        
        self.status_label = ttk.Label(
            self.progress_frame, 
            text="Bereit zum Konvertieren",
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X, pady=(5, 0))
        
        # Konvertierungs-Button
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.convert_btn = ttk.Button(
            button_frame, 
            text="Konvertieren", 
            command=self.start_conversion
        )
        self.convert_btn.pack(pady=5, ipadx=20, ipady=5)
    
    def create_log_tab(self, parent):
        log_frame = ttk.Frame(parent)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD,
            state=tk.DISABLED
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        btn_frame = ttk.Frame(log_frame)
        btn_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(
            btn_frame, 
            text="Protokoll exportieren", 
            command=self.export_log
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            btn_frame, 
            text="Protokoll löschen", 
            command=self.clear_log
        ).pack(side=tk.LEFT, padx=5)
    
    def update_threshold_label(self, *args):
        self.threshold_label.config(text=str(self.threshold_value.get()))
    
    def update_compression_label(self, *args):
        self.compression_label.config(text=str(self.compression_var.get()))
    
    def browse_input(self):
        file_paths = filedialog.askopenfilenames(
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        if file_paths:
            self.input_files = list(file_paths)
            self.input_listbox.delete(0, tk.END)
            for file_path in self.input_files:
                self.input_listbox.insert(tk.END, os.path.basename(file_path))
    
    def clear_input(self):
        self.input_files = []
        self.input_listbox.delete(0, tk.END)
    
    def browse_output_dir(self):
        dir_path = filedialog.askdirectory()
        if dir_path:
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, dir_path)
    
    def save_app_settings(self):
        self.settings['threshold'] = self.threshold_value.get()
        self.settings['dpi'] = int(self.dpi_value.get())
        self.settings['open_after'] = self.open_after_var.get()
        self.settings['output_dir'] = self.output_dir_entry.get().strip()
        self.settings['output_suffix'] = self.suffix_entry.get().strip()
        self.settings['page_range'] = self.page_range_entry.get().strip()
        self.settings['mode'] = self.mode_var.get()
        self.settings['overwrite'] = self.overwrite_var.get()
        self.settings['compression'] = self.compression_var.get()
        
        save_settings(self.settings)
    
    def log_message(self, message):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        logger.info(message)
    
    def update_progress(self, percent, message):
        if percent >= 0:
            self.progress_bar['value'] = percent
        
        self.status_label.config(text=message)
        self.log_message(message)
        self.root.update_idletasks()
    
    def start_conversion(self):
        if not self.input_files:
            messagebox.showwarning("Keine Dateien", "Bitte wählen Sie mindestens eine PDF-Datei aus.")
            return
        
        # Einstellungen speichern
        self.save_app_settings()
        
        self.current_file_index = 0
        self.total_files = len(self.input_files)
        
        # Log zurücksetzen
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.status_label.config(text="Konvertierung wird gestartet...")
        self.convert_btn.config(state=tk.DISABLED)
        
        # Starte Konvertierung
        self.convert_next_file()
    
    def convert_next_file(self):
        if self.current_file_index >= self.total_files:
            self.update_progress(100, "✅ Alle Konvertierungen abgeschlossen!")
            self.convert_btn.config(state=tk.NORMAL)
            return
        
        input_path = self.input_files[self.current_file_index]
        output_path = self.generate_output_path(input_path)
        
        # Überschreiben prüfen
        if os.path.exists(output_path):
            overwrite = self.settings['overwrite']
            if overwrite == 'ask':
                response = messagebox.askyesno(
                    "Datei existiert",
                    f"Die Ausgabedatei '{os.path.basename(output_path)}' existiert bereits. Überschreiben?"
                )
                if not response:
                    self.log_message(f"Überspringen: {os.path.basename(input_path)}")
                    self.current_file_index += 1
                    self.convert_next_file()
                    return
            elif overwrite == 'skip':
                self.log_message(f"Überspringen: {os.path.basename(input_path)}")
                self.current_file_index += 1
                self.convert_next_file()
                return
        
        # Seitenbereich analysieren
        page_range_str = self.page_range_entry.get().strip()
        page_range = None
        if page_range_str:
            try:
                with fitz.open(input_path) as doc:
                    total_pages = len(doc)
                    page_range = parse_page_range(page_range_str, total_pages)
            except Exception as e:
                self.log_message(f"Fehler beim Analysieren des Seitenbereichs: {str(e)}")
                page_range = None
        
        # Starte Konvertierung im Hintergrund
        threading.Thread(
            target=self.run_conversion,
            args=(input_path, output_path, page_range),
            daemon=True
        ).start()
    
    def generate_output_path(self, input_path):
        output_dir = self.output_dir_entry.get().strip() or os.path.dirname(input_path)
        filename = os.path.basename(input_path)
        base, ext = os.path.splitext(filename)
        
        # Suffix hinzufügen
        suffix = self.suffix_entry.get().strip()
        if suffix:
            base += suffix
        
        # Zeitstempel hinzufügen
        if self.timestamp_var.get():
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            base += f"_{timestamp}"
        
        return os.path.join(output_dir, base + ext)
    
    def run_conversion(self, input_path, output_path, page_range):
        def progress_callback(percent, message):
            self.root.after(10, lambda: self.update_progress(percent, message))
        
        success = convert_pdf_to_bw(
            input_path,
            output_path,
            threshold=self.threshold_value.get(),
            dpi=int(self.dpi_value.get()),
            progress_callback=progress_callback,
            page_range=page_range,
            mode=self.mode_var.get(),
            compression=self.compression_var.get()
        )
        
        if success and self.open_after_var.get():
            self.root.after(100, lambda: self.open_file(output_path))
        
        self.current_file_index += 1
        self.root.after(10, self.convert_next_file)
    
    def open_file(self, path):
        try:
            if platform.system() == 'Darwin':  # macOS
                os.system(f'open "{path}"')
            elif platform.system() == 'Windows':  # Windows
                os.startfile(path)
            else:  # Linux und andere
                webbrowser.open(path)
        except Exception as e:
            self.log_message(f"❌ Fehler beim Öffnen der Datei: {str(e)}")
    
    def export_log(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Textdateien", "*.txt"), ("Alle Dateien", "*.*")]
        )
        if file_path:
            try:
                with open(file_path, "w") as f:
                    f.write(self.log_text.get("1.0", tk.END))
                self.log_message(f"Protokoll exportiert: {file_path}")
            except Exception as e:
                messagebox.showerror("Fehler", f"Protokoll konnte nicht exportiert werden: {str(e)}")
    
    def clear_log(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

if __name__ == "__main__":  
    root = tk.Tk()
    app = PDFConverterApp(root)
    root.mainloop()

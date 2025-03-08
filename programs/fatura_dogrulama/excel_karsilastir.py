import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime
from tkinter import *
import subprocess
import platform

class ExcelKarsilastir:
    def __init__(self, root):
        self.root = root
        
        # Tam ekran olarak ayarla
        self.root.state('zoomed')  # Windows için tam ekran
        
        # Minimum boyut
        self.root.minsize(900, 600)
        
        # Modern renk paleti
        self.primary_color = "#2563eb"  # Mavi
        self.primary_light = "#3b82f6"  # Açık mavi
        self.primary_dark = "#1d4ed8"   # Koyu mavi
        self.secondary_color = "#0f172a"  # Koyu lacivert
        self.accent_color = "#f97316"  # Turuncu
        self.bg_color = "#f8fafc"  # Çok açık gri
        self.card_bg = "#ffffff"  # Beyaz
        self.text_color = "#1e293b"  # Koyu gri
        self.text_light = "#64748b"  # Açık gri
        
        # Style ayarları
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Genel stil ayarları
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('Card.TFrame', background=self.card_bg)
        
        self.style.configure('TLabel', 
                            background=self.bg_color, 
                            foreground=self.text_color, 
                            font=('Segoe UI', 10))
        
        self.style.configure('Card.TLabel', 
                            background=self.card_bg, 
                            foreground=self.text_color, 
                            font=('Segoe UI', 10))
        
        self.style.configure('Header.TLabel', 
                            font=('Segoe UI', 22, 'bold'), 
                            foreground=self.secondary_color,
                            background=self.bg_color)
        
        self.style.configure('Subheader.TLabel', 
                            font=('Segoe UI', 14), 
                            foreground=self.text_color,
                            background=self.bg_color)
        
        # LabelFrame stil ayarları
        self.style.configure('TLabelframe', 
                            background=self.bg_color,
                            foreground=self.text_color,
                            font=('Segoe UI', 11, 'bold'))
        
        self.style.configure('TLabelframe.Label', 
                            background=self.bg_color,
                            foreground=self.primary_color,
                            font=('Segoe UI', 11, 'bold'))
        
        # Buton stilleri
        self.style.configure('Primary.TButton', 
                            font=('Segoe UI', 11, 'bold'),
                            padding=(15, 10))
        
        self.style.map('Primary.TButton',
                      background=[('active', self.primary_light), 
                                 ('pressed', self.primary_dark),
                                 ('!active', self.primary_color)],
                      foreground=[('active', 'white'), 
                                 ('pressed', 'white'),
                                 ('!active', 'white')])
        
        self.style.configure('Secondary.TButton', 
                            font=('Segoe UI', 11),
                            padding=(15, 8))
        
        self.style.map('Secondary.TButton',
                      background=[('active', '#f1f5f9'), 
                                 ('pressed', '#e2e8f0'),
                                 ('!active', '#f8fafc')],
                      foreground=[('active', self.primary_color), 
                                 ('pressed', self.primary_dark),
                                 ('!active', self.primary_color)])
        
        # Dosya yolları
        self.gelen_fatura_path = tk.StringVar()
        self.islenmis_fatura_path = tk.StringVar()
        
        # Seçili sütunlar ve seçim sırası
        self.selected_columns_gelen = []
        self.selected_columns_islenmis = []
        self.selection_order_gelen = []
        self.selection_order_islenmis = []
        
        # Dosya yükleme durumu
        self.gelen_loaded = False
        self.islenmis_loaded = False
        
        # Arayüz oluşturma
        self.create_widgets()
    
    def create_widgets(self):
        # Ana frame
        main_frame = ttk.Frame(self.root, style='TFrame', padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scroll için canvas oluştur
        canvas = tk.Canvas(main_frame, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        
        # Scrollable frame
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Canvas içine frame yerleştir
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", width=canvas.winfo_reqwidth())
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Canvas ve scrollbar'ı yerleştir
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mouse wheel ile scroll
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Canvas genişliğini pencere genişliğine ayarla
        def _configure_canvas(event):
            canvas_width = event.width
            canvas.itemconfig(canvas.find_withtag("all")[0], width=canvas_width)
        
        canvas.bind("<Configure>", _configure_canvas)
        
        # İçerik için ana frame
        content_frame = ttk.Frame(scrollable_frame, style='TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Başlık
        header_frame = ttk.Frame(content_frame, style='TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Dosya seçme bölümü
        file_frame = ttk.LabelFrame(content_frame, text="Dosya Seçimi", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        file_frame.columnconfigure(1, weight=1)
        
        # Gelen fatura dosyası
        ttk.Label(file_frame, text="Gelen Fatura:").grid(row=0, column=0, sticky="w", padx=(0, 10), pady=10)
        ttk.Entry(file_frame, textvariable=self.gelen_fatura_path, width=50).grid(row=0, column=1, sticky="ew", pady=10)
        ttk.Button(file_frame, text="Dosya Seç", command=self.select_gelen_fatura, style='Secondary.TButton').grid(row=0, column=2, padx=10, pady=10)
        
        # İşlenmiş fatura dosyası
        ttk.Label(file_frame, text="İşlenmiş Fatura:").grid(row=1, column=0, sticky="w", padx=(0, 10), pady=10)
        ttk.Entry(file_frame, textvariable=self.islenmis_fatura_path, width=50).grid(row=1, column=1, sticky="ew", pady=10)
        ttk.Button(file_frame, text="Dosya Seç", command=self.select_islenmis_fatura, style='Secondary.TButton').grid(row=1, column=2, padx=10, pady=10)
        
        # Seçili sütunları gösterme alanı - Yukarı taşındı
        selected_frame = ttk.Frame(content_frame, style='TFrame', padding="10")
        selected_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(selected_frame, text="Seçili Sütunlar:", style='Subheader.TLabel').pack(anchor="w", pady=(0, 5))
        
        self.selected_text = Text(selected_frame, height=4, font=('Segoe UI', 10), wrap=tk.WORD,
                                bg=self.card_bg, fg=self.text_color)
        self.selected_text.pack(fill=tk.X)
        self.selected_text.insert(tk.END, "Henüz sütun seçilmedi")
        self.selected_text.config(state=tk.DISABLED)
        
        # Karşılaştırma ayarları
        settings_frame = ttk.LabelFrame(content_frame, text="Karşılaştırma Ayarları", padding="10")
        settings_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Sütun seçim alanı
        columns_frame = ttk.Frame(settings_frame, style='TFrame')
        columns_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        columns_frame.columnconfigure(0, weight=1)
        columns_frame.columnconfigure(1, weight=1)
        
        # Gelen fatura sütunları
        gelen_frame = ttk.LabelFrame(columns_frame, text="Gelen Fatura Sütunları", padding="10")
        gelen_frame.grid(row=0, column=0, padx=5, sticky=tk.NSEW)
        
        self.gelen_listbox = Listbox(gelen_frame, selectmode=MULTIPLE, height=10,
                                   exportselection=0, font=('Segoe UI', 10),
                                   bg=self.card_bg, fg=self.text_color)
        gelen_scrollbar = ttk.Scrollbar(gelen_frame, orient="vertical", command=self.gelen_listbox.yview)
        self.gelen_listbox.configure(yscrollcommand=gelen_scrollbar.set)
        self.gelen_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        gelen_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.gelen_listbox.bind('<<ListboxSelect>>', self.on_select_gelen)
        
        # İşlenmiş fatura sütunları
        islenmis_frame = ttk.LabelFrame(columns_frame, text="İşlenmiş Fatura Sütunları", padding="10")
        islenmis_frame.grid(row=0, column=1, padx=5, sticky=tk.NSEW)
        
        self.islenmis_listbox = Listbox(islenmis_frame, selectmode=MULTIPLE, height=10,
                                      exportselection=0, font=('Segoe UI', 10),
                                      bg=self.card_bg, fg=self.text_color)
        islenmis_scrollbar = ttk.Scrollbar(islenmis_frame, orient="vertical", command=self.islenmis_listbox.yview)
        self.islenmis_listbox.configure(yscrollcommand=islenmis_scrollbar.set)
        self.islenmis_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        islenmis_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.islenmis_listbox.bind('<<ListboxSelect>>', self.on_select_islenmis)
        
        # Butonlar ve durum çubuğu için alt frame
        bottom_frame = ttk.Frame(content_frame, style='TFrame', padding=(0, 20, 0, 10))
        bottom_frame.pack(fill=tk.X)
        
        # Butonlar
        button_frame = ttk.Frame(bottom_frame, style='TFrame')
        button_frame.pack(fill=tk.X)
        
        # Butonları ortalamak için pack kullanımı
        ttk.Frame(button_frame, style='TFrame').pack(side=tk.LEFT, expand=True)
        
        # Rapor Oluştur butonu
        report_button = ttk.Button(
            button_frame, 
            text="Rapor Oluştur", 
            command=self.create_report, 
            style='Primary.TButton',
            width=20
        )
        report_button.pack(side=tk.LEFT, padx=10)
        
        # Çıkış butonu
        exit_button = ttk.Button(
            button_frame, 
            text="Çıkış", 
            command=self.root.quit, 
            style='Secondary.TButton',
            width=15
        )
        exit_button.pack(side=tk.LEFT, padx=10)
        
        ttk.Frame(button_frame, style='TFrame').pack(side=tk.LEFT, expand=True)
        
        # Durum çubuğu
        status_frame = ttk.Frame(bottom_frame, style='TFrame')
        status_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.status_var = tk.StringVar()
        self.status_var.set("Hazır")
        ttk.Label(status_frame, textvariable=self.status_var, foreground=self.text_light).pack(side=tk.LEFT)
    
    def on_select_gelen(self, evt):
        current_selection = list(self.gelen_listbox.curselection())
        
        # Yeni seçilen öğeyi bul
        if len(current_selection) > len(self.selection_order_gelen):
            # Yeni bir öğe seçildi
            new_items = set(current_selection) - set(self.selection_order_gelen)
            for item in new_items:
                if item not in self.selection_order_gelen:
                    self.selection_order_gelen.append(item)
        else:
            # Bir öğe seçimden kaldırıldı
            removed_items = set(self.selection_order_gelen) - set(current_selection)
            for item in removed_items:
                self.selection_order_gelen.remove(item)
        
        # Seçili sütunları seçim sırasına göre güncelle
        self.selected_columns_gelen = [self.gelen_listbox.get(i) for i in self.selection_order_gelen]
        self.update_selected_columns_label()
    
    def on_select_islenmis(self, evt):
        current_selection = list(self.islenmis_listbox.curselection())
        
        # Yeni seçilen öğeyi bul
        if len(current_selection) > len(self.selection_order_islenmis):
            # Yeni bir öğe seçildi
            new_items = set(current_selection) - set(self.selection_order_islenmis)
            for item in new_items:
                if item not in self.selection_order_islenmis:
                    self.selection_order_islenmis.append(item)
        else:
            # Bir öğe seçimden kaldırıldı
            removed_items = set(self.selection_order_islenmis) - set(current_selection)
            for item in removed_items:
                self.selection_order_islenmis.remove(item)
        
        # Seçili sütunları seçim sırasına göre güncelle
        self.selected_columns_islenmis = [self.islenmis_listbox.get(i) for i in self.selection_order_islenmis]
        self.update_selected_columns_label()
    
    def update_selected_columns_label(self):
        self.selected_text.config(state=tk.NORMAL)
        self.selected_text.delete(1.0, tk.END)
        
        if not self.selection_order_gelen and not self.selection_order_islenmis:
            self.selected_text.insert(tk.END, "Henüz sütun seçilmedi")
        else:
            text = "Eşleştirilen Sütunlar:\n"
            for i in range(max(len(self.selection_order_gelen), len(self.selection_order_islenmis))):
                gelen = self.gelen_listbox.get(self.selection_order_gelen[i]) if i < len(self.selection_order_gelen) else ""
                islenmis = self.islenmis_listbox.get(self.selection_order_islenmis[i]) if i < len(self.selection_order_islenmis) else ""
                text += f"{i+1}. {gelen} ↔ {islenmis}\n"
            self.selected_text.insert(tk.END, text)
        
        self.selected_text.config(state=tk.DISABLED)
    
    def select_gelen_fatura(self):
        filename = filedialog.askopenfilename(
            title="Gelen Fatura Listesi Seç",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        if filename:
            self.gelen_fatura_path.set(filename)
            self.gelen_loaded = True
            if self.islenmis_loaded:
                self.load_files()
    
    def select_islenmis_fatura(self):
        filename = filedialog.askopenfilename(
            title="İşlenmiş Fatura Listesi Seç",
            filetypes=[("Excel Dosyaları", "*.xlsx *.xls")]
        )
        if filename:
            self.islenmis_fatura_path.set(filename)
            self.islenmis_loaded = True
            if self.gelen_loaded:
                self.load_files()
    
    def load_files(self):
        try:
            # Excel dosyalarını oku
            self.status_var.set("Dosyalar yükleniyor...")
            self.root.update_idletasks()
            
            # Sayısal sütunları string olarak oku
            self.gelen_df = pd.read_excel(self.gelen_fatura_path.get(), dtype=str)
            self.islenmis_df = pd.read_excel(self.islenmis_fatura_path.get(), dtype=str)
            
            # NaN değerleri boş string ile değiştir
            self.gelen_df = self.gelen_df.fillna('')
            self.islenmis_df = self.islenmis_df.fillna('')
            
            # Tüm değerleri string'e çevir ve temizle
            for df in [self.gelen_df, self.islenmis_df]:
                for column in df.columns:
                    df[column] = df[column].astype(str).apply(lambda x: x.strip().rstrip('.0') if x.endswith('.0') else x.strip())
            
            # Listboxları temizle
            self.gelen_listbox.delete(0, tk.END)
            self.islenmis_listbox.delete(0, tk.END)
            
            # Sütunları listboxlara ekle
            for col in self.gelen_df.columns:
                self.gelen_listbox.insert(tk.END, col)
            
            for col in self.islenmis_df.columns:
                self.islenmis_listbox.insert(tk.END, col)
            
            self.status_var.set(f"Dosyalar yüklendi. Gelen: {len(self.gelen_df)} satır, İşlenmiş: {len(self.islenmis_df)} satır.")
            
        except Exception as e:
            self.status_var.set("Hata: Dosyalar yüklenemedi!")
            messagebox.showerror("Hata", f"Dosyalar yüklenirken bir hata oluştu:\n{str(e)}")
            self.gelen_loaded = False
            self.islenmis_loaded = False
    
    def compare_files(self):
        if not hasattr(self, 'gelen_df') or not hasattr(self, 'islenmis_df'):
            messagebox.showerror("Hata", "Lütfen önce dosyaları yükleyin!")
            return
        
        if len(self.selected_columns_gelen) < 4 or len(self.selected_columns_islenmis) < 4:
            messagebox.showerror("Hata", "Lütfen her iki dosya için de en az 4 sütun seçin!")
            return
        
        try:
            self.status_var.set("Karşılaştırma yapılıyor...")
            self.root.update_idletasks()
            
            # İlk 4 sütunu al
            primary_columns_gelen = self.selected_columns_gelen[:4]
            primary_columns_islenmis = self.selected_columns_islenmis[:4]
            
            # 3'lü kombinasyonlar
            combinations = [
                ([0,1,2], [3]),  # ABC-D
                ([0,1,3], [2]),  # ABD-C
                ([0,2,3], [1]),  # ACD-B
                ([1,2,3], [0])   # BCD-A
            ]
            
            # Eşleşme sayacı
            partial_match_count = 0  # 3 sütun eşleşip 1 sütun eşleşmeyenler
            full_match_count = 0     # 4 sütun da eşleşenler
            
            # Her bir gelen fatura için
            for idx_g, row_g in self.gelen_df.iterrows():
                # Her bir işlenmiş fatura ile karşılaştır
                for idx_i, row_i in self.islenmis_df.iterrows():
                    # Önce tam eşleşme kontrolü
                    full_match = True
                    for i in range(4):
                        val_g = str(row_g[primary_columns_gelen[i]]).strip()
                        val_i = str(row_i[primary_columns_islenmis[i]]).strip()
                        if val_g != val_i:
                            full_match = False
                            break
                    
                    if full_match:
                        full_match_count += 1
                        continue
                    
                    # 3'lü kombinasyonlar için kontrol
                    for match_indices, diff_indices in combinations:
                        # Eşleşen sütunları kontrol et
                        all_match = True
                        for idx in match_indices:
                            val_g = str(row_g[primary_columns_gelen[idx]]).strip()
                            val_i = str(row_i[primary_columns_islenmis[idx]]).strip()
                            if val_g != val_i:
                                all_match = False
                                break
                        
                        # Eğer 3 sütun eşleşiyorsa ve 4. sütun eşleşmiyorsa
                        if all_match:
                            for diff_idx in diff_indices:
                                val_g = str(row_g[primary_columns_gelen[diff_idx]]).strip()
                                val_i = str(row_i[primary_columns_islenmis[diff_idx]]).strip()
                                if val_g != val_i:
                                    partial_match_count += 1
            
            # Sonuç mesajı
            self.status_var.set(f"Karşılaştırma tamamlandı. Tam Eşleşen: {full_match_count}, " +
                              f"Kısmi Eşleşen (3/4): {partial_match_count}")
            
            # Bilgi etiketini güncelle
            self.selected_text.config(state=tk.NORMAL)
            self.selected_text.delete(1.0, tk.END)
            self.selected_text.insert(tk.END, f"✓ Karşılaştırma tamamlandı! Tam Eşleşen: {full_match_count}, Kısmi Eşleşen: {partial_match_count}")
            self.selected_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.status_var.set("Hata: Karşılaştırma yapılamadı!")
            messagebox.showerror("Hata", f"Karşılaştırma sırasında bir hata oluştu:\n{str(e)}")
    
    def is_date(self, s):
        """Verilen string'in tarih formatında olup olmadığını kontrol eder"""
        try:
            pd.to_datetime(s)
            return True
        except:
            return False

    def is_number(self, s):
        """Verilen string'in sayı olup olmadığını kontrol eder"""
        try:
            # Virgülü noktaya çevir
            s = s.replace(',', '.')
            float(s)
            return True
        except ValueError:
            return False

    def check_numeric_difference(self, val1, val2):
        """İki sayısal değer arasındaki farkın 0.02'den küçük olup olmadığını kontrol eder"""
        try:
            # Virgülleri nokta ile değiştir
            num1 = float(val1.replace(',', '.'))
            num2 = float(val2.replace(',', '.'))
            return abs(num1 - num2) <= 0.02
        except:
            return False

    def find_currency_column(self, df):
        """Döviz cinsi sütununu bulur"""
        # İsme göre ara
        for col in df.columns:
            if 'döviz' in col.lower() or 'doviz' in col.lower():
                return col
        
        # İçeriğe göre ara (sadece TRY olan sütun)
        for col in df.columns:
            unique_values = df[col].unique()
            if len(unique_values) == 1 and 'TRY' in unique_values:
                return col
        
        return None

    def check_currency_mismatch(self, row, currency_col):
        """Döviz değerinin TRY olup olmadığını kontrol eder"""
        if currency_col:
            currency_value = str(row[currency_col]).strip().upper()
            return currency_value != 'TRY'
        return False

    def create_report(self):
        if not hasattr(self, 'gelen_df') or not hasattr(self, 'islenmis_df'):
            messagebox.showerror("Hata", "Lütfen önce dosyaları yükleyin!")
            return
        
        if len(self.selected_columns_gelen) < 4 or len(self.selected_columns_islenmis) < 4:
            messagebox.showerror("Hata", "Lütfen her iki dosya için de en az 4 sütun seçin!")
            return
        
        try:
            self.status_var.set("Rapor oluşturuluyor...")
            self.root.update_idletasks()
            
            # Masaüstü yolunu al
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            
            # Rapor klasörü oluştur
            report_dir = os.path.join(desktop_path, "Aylık Fatura Doğrulama Raporları")
            if not os.path.exists(report_dir):
                os.makedirs(report_dir)
            
            # Rapor dosya adı
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            report_file = os.path.join(report_dir, f"Fatura_Karsilastirma_Raporu_{timestamp}.xlsx")

            
            # İlk 4 sütunu al
            primary_columns_gelen = self.selected_columns_gelen[:4]
            primary_columns_islenmis = self.selected_columns_islenmis[:4]
            
            # Döviz sütununu bul
            currency_column = self.find_currency_column(self.gelen_df)
            
            # Excel writer oluştur
            with pd.ExcelWriter(report_file, engine='xlsxwriter') as writer:
                workbook = writer.book
                
                # Kırmızı format
                red_format = workbook.add_format({
                    'bg_color': '#FFC7CE',
                    'font_color': '#9C0006'
                })
                
                # Header format
                header_format = workbook.add_format({
                    'bold': True,
                    'bg_color': '#FFA07A',
                    'border': 1
                })

                # Sadece gelen faturalarda olan kayıtlar için format
                gelen_only_format = workbook.add_format({
                    'bg_color': '#E6B8B7',
                    'font_color': '#000000'
                })

                # Sadece işlenmiş faturalarda olan kayıtlar için format
                islenmis_only_format = workbook.add_format({
                    'bg_color': '#B8CCE4',
                    'font_color': '#000000'
                })
                
                mismatch_details = []
                gelen_only_records = []
                islenmis_only_records = []
                
                # Her bir gelen fatura için
                for idx_g, row_g in self.gelen_df.iterrows():
                    found_match = False
                    # Her bir işlenmiş fatura ile karşılaştır
                    for idx_i, row_i in self.islenmis_df.iterrows():
                        # 3'lü kombinasyonlar için kontrol
                        combinations = [
                            ([0,1,2], [3]),  # ABC-D
                            ([0,1,3], [2]),  # ABD-C
                            ([0,2,3], [1]),  # ACD-B
                            ([1,2,3], [0])   # BCD-A
                        ]
                        
                        for match_indices, diff_indices in combinations:
                            # Eşleşen sütunları kontrol et
                            all_match = True
                            for idx in match_indices:
                                val_g = str(row_g[primary_columns_gelen[idx]]).strip()
                                val_i = str(row_i[primary_columns_islenmis[idx]]).strip()
                                if val_g != val_i:
                                    all_match = False
                                    break
                            
                            if all_match:
                                found_match = True
                                for diff_idx in diff_indices:
                                    val_g = str(row_g[primary_columns_gelen[diff_idx]]).strip()
                                    val_i = str(row_i[primary_columns_islenmis[diff_idx]]).strip()
                                    
                                    if val_g != val_i:
                                        # Sayısal değer kontrolü
                                        if self.is_number(val_g) and self.is_number(val_i):
                                            if self.check_numeric_difference(val_g, val_i):
                                                continue  # Fark 0.02'den küçükse raporda gösterme
                                        
                                        details = {}
                                        
                                        # İlk seçilen sütun farklı ise
                                        if diff_idx == 0:
                                            details = {
                                                primary_columns_gelen[0]: str(row_g[primary_columns_gelen[0]]).strip(),  # EFatura ID'yi boş bırakmak yerine göster
                                            }
                                            # Diğer sütunları da doldur
                                            for i in range(1, 4):
                                                details[primary_columns_gelen[i]] = str(row_g[primary_columns_gelen[i]]).strip()
                                            
                                            details.update({
                                                'Farklı Sütun': f"{primary_columns_gelen[diff_idx]} - {primary_columns_islenmis[diff_idx]}",
                                                'Gelen Değer': val_g,
                                                'İşlenmiş Değer': val_i
                                            })
                                            
                                            # Döviz kontrolü
                                            if self.check_currency_mismatch(row_g, currency_column):
                                                details['Sebebi'] = 'Döviz cinsinin farklılığından ötürü.'
                                        else:
                                            # Normal durum - tüm sütunları doldur
                                            details = {
                                                primary_columns_gelen[0]: str(row_g[primary_columns_gelen[0]]).strip(),
                                            }
                                            # Diğer sütunları da doldur
                                            for i in range(1, 4):
                                                details[primary_columns_gelen[i]] = str(row_g[primary_columns_gelen[i]]).strip()
                                            
                                            details.update({
                                                'Farklı Sütun': f"{primary_columns_gelen[diff_idx]} - {primary_columns_islenmis[diff_idx]}",
                                                'Gelen Değer': val_g,
                                                'İşlenmiş Değer': val_i
                                            })
                                            
                                            # Döviz kontrolü
                                            if self.check_currency_mismatch(row_g, currency_column):
                                                details['Sebebi'] = 'Döviz cinsinin farklılığından ötürü.'
                                        
                                        mismatch_details.append(details)

                    # Eğer hiç eşleşme bulunamadıysa, sadece gelen faturalarda var demektir
                    if not found_match:
                        record_dict = {}
                        # Tüm sütunları al
                        for col in self.gelen_df.columns:
                            record_dict[col] = str(row_g[col]).strip()
                        record_dict['Durum'] = 'Sadece Gelen Faturalarda'
                        gelen_only_records.append(record_dict)

                # İşlenmiş faturalarda olup gelen faturalarda olmayanları bul
                for idx_i, row_i in self.islenmis_df.iterrows():
                    found_in_gelen = False
                    for idx_g, row_g in self.gelen_df.iterrows():
                        # İlk iki sütuna göre karşılaştır
                        if (str(row_i[primary_columns_islenmis[0]]).strip() == str(row_g[primary_columns_gelen[0]]).strip() and
                            str(row_i[primary_columns_islenmis[1]]).strip() == str(row_g[primary_columns_gelen[1]]).strip()):
                            found_in_gelen = True
                            break
                    
                    if not found_in_gelen:
                        record_dict = {}
                        # Tüm sütunları al
                        for col in self.islenmis_df.columns:
                            record_dict[col] = str(row_i[col]).strip()
                        record_dict['Durum'] = 'Sadece İşlenmiş Faturalarda'
                        islenmis_only_records.append(record_dict)
                
                # Eşleşmeyen kayıtları yaz
                if mismatch_details:
                    # Yeni sütun yapısını oluştur
                    columns = primary_columns_gelen[:4]  # İlk 4 seçili sütun
                    columns.extend(['Farklı Sütun', 'Gelen Değer', 'İşlenmiş Değer', 'Sebebi', 'SERİ NO', 'SIRA NO', 'CİRO CARİ İSMİ'])
                    
                    # Boş bir DataFrame oluştur
                    mismatch_df = pd.DataFrame(columns=columns)
                    
                    # Her bir eşleşmeyen kayıt için
                    for detail in mismatch_details:
                        row_data = {}
                        # İlk 4 sütunu doldur
                        for i, col in enumerate(primary_columns_gelen[:4]):
                            if i == 0:  # İlk sütun için özel işlem
                                row_data[col] = detail.get(col, '')
                            else:
                                # Eşleşmeyen sütun bu mu kontrol et
                                if detail['Farklı Sütun'].startswith(col):
                                    row_data[col] = detail['Gelen Değer']
                                else:
                                    # İşlenmiş faturadan değeri bul
                                    for idx_i, row_i in self.islenmis_df.iterrows():
                                        if str(row_i[primary_columns_islenmis[0]]) == str(detail.get(primary_columns_gelen[0], '')):
                                            row_data[col] = str(row_i[primary_columns_islenmis[i]]).strip()
                                            break
                        
                        # Diğer sütunları doldur
                        row_data['Farklı Sütun'] = detail['Farklı Sütun']
                        row_data['Gelen Değer'] = detail['Gelen Değer']
                        row_data['İşlenmiş Değer'] = detail['İşlenmiş Değer']
                        row_data['Sebebi'] = detail.get('Sebebi', '')
                        
                        # SERİ, SIRA NO ve CİRO CARİ İSMİ'ni bul
                        for idx_i, row_i in self.islenmis_df.iterrows():
                            if str(row_i[primary_columns_islenmis[0]]) == str(detail.get(primary_columns_gelen[0], '')):
                                seri_col = next((col for col in self.islenmis_df.columns if 'SERİ' in col.upper()), None)
                                sira_col = next((col for col in self.islenmis_df.columns if 'SIRA' in col.upper()), None)
                                cari_col = next((col for col in self.islenmis_df.columns if 'CİRO CARİ İSMİ' in col.upper()), None)
                                
                                row_data['SERİ NO'] = str(row_i[seri_col]).strip() if seri_col else ''
                                row_data['SIRA NO'] = str(row_i[sira_col]).strip() if sira_col else ''
                                row_data['CİRO CARİ İSMİ'] = str(row_i[cari_col]).strip() if cari_col else ''
                                break
                        
                        # DataFrame'e ekle
                        mismatch_df = pd.concat([mismatch_df, pd.DataFrame([row_data])], ignore_index=True)
                    
                    # DataFrame'i sırala ve Excel'e yaz
                    first_column = mismatch_df.columns[0]
                    mismatch_df = mismatch_df.sort_values(by=first_column, na_position='last')
                    mismatch_df.to_excel(writer, sheet_name='Eşleşmeyen Detaylar', index=False)
                    
                    worksheet = writer.sheets['Eşleşmeyen Detaylar']
                    
                    # Sütun genişliklerini ayarla
                    for col_num, column in enumerate(mismatch_df.columns):
                        column_width = max(len(str(val)) for val in mismatch_df[column])
                        column_width = max(column_width, len(column)) + 2
                        worksheet.set_column(col_num, col_num, min(column_width, 30))
                    
                    # Header formatını uygula
                    for col_num, value in enumerate(mismatch_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    # Eşleşmeyen değerleri kırmızı yap
                    for row_num in range(len(mismatch_df)):
                        for col_num, column in enumerate(mismatch_df.columns):
                            cell_value = mismatch_df.iloc[row_num, col_num]
                            
                            # Eşleşmeyen sütunu bul
                            farkli_sutun = mismatch_df.iloc[row_num]['Farklı Sütun'].split(' - ')[0]
                            
                            # Eğer bu sütun eşleşmeyen sütun ise kırmızı yap
                            if column == farkli_sutun:
                                worksheet.write(row_num + 1, col_num, cell_value, red_format)
                            else:
                                worksheet.write(row_num + 1, col_num, cell_value)

                # Sadece gelen faturalarda olan kayıtları yaz
                if gelen_only_records:
                    gelen_only_df = pd.DataFrame(gelen_only_records)
                    gelen_only_df.to_excel(writer, sheet_name='Sadece Gelen Faturalar', index=False)
                    
                    worksheet = writer.sheets['Sadece Gelen Faturalar']
                    # Tüm sütunlar için genişlik ayarla
                    for col_num in range(len(gelen_only_df.columns)):
                        worksheet.set_column(col_num, col_num, 30)
                    
                    for col_num, value in enumerate(gelen_only_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    for row_num in range(len(gelen_only_df)):
                        for col_num in range(len(gelen_only_df.columns)):
                            worksheet.write(row_num + 1, col_num, gelen_only_df.iloc[row_num, col_num], gelen_only_format)

                # Sadece işlenmiş faturalarda olan kayıtları yaz
                if islenmis_only_records:
                    islenmis_only_df = pd.DataFrame(islenmis_only_records)
                    islenmis_only_df.to_excel(writer, sheet_name='Sadece İşlenmiş Faturalar', index=False)
                    
                    worksheet = writer.sheets['Sadece İşlenmiş Faturalar']
                    # Tüm sütunlar için genişlik ayarla
                    for col_num in range(len(islenmis_only_df.columns)):
                        worksheet.set_column(col_num, col_num, 30)
                    
                    for col_num, value in enumerate(islenmis_only_df.columns.values):
                        worksheet.write(0, col_num, value, header_format)
                    
                    for row_num in range(len(islenmis_only_df)):
                        for col_num in range(len(islenmis_only_df.columns)):
                            worksheet.write(row_num + 1, col_num, islenmis_only_df.iloc[row_num, col_num], islenmis_only_format)
            
            self.status_var.set(f"Rapor oluşturuldu: {report_file}")
            self.selected_text.config(state=tk.NORMAL)
            self.selected_text.delete(1.0, tk.END)
            self.selected_text.insert(tk.END, "✓ Rapor başarıyla oluşturuldu!")
            self.selected_text.config(state=tk.DISABLED)
            
            # Raporu otomatik aç
            try:
                if platform.system() == 'Windows':
                    os.startfile(report_file)
                elif platform.system() == 'Darwin':  # macOS
                    subprocess.run(['open', report_file])
                else:  # Linux
                    subprocess.run(['xdg-open', report_file])
            except Exception as e:
                self.selected_text.config(state=tk.NORMAL)
                self.selected_text.delete(1.0, tk.END)
                self.selected_text.insert(tk.END, "✓ Rapor oluşturuldu fakat otomatik açılamadı")
                self.selected_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.status_var.set("Hata: Rapor oluşturulamadı!")
            messagebox.showerror("Hata", f"Rapor oluşturulurken bir hata oluştu:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelKarsilastir(root)
    root.mainloop() 
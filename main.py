import tkinter as tk
from tkinter import ttk, messagebox
import os
import sys
import importlib.util
import subprocess
import platform

class DokucuogluYazilim:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Dokucuoglu Yazılım - Program Merkezi")
        
        # Tam ekran olarak ayarla
        self.root.state('zoomed')  # Windows için tam ekran
        
        # Minimum boyut
        self.root.minsize(900, 600)
        
        # Tema ve stil ayarları
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
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
        
        # Stil konfigürasyonları
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
        
        self.style.configure('Title.TLabel', 
                            font=('Segoe UI', 16, 'bold'), 
                            foreground=self.primary_color,
                            background=self.card_bg)
        
        self.style.configure('Description.TLabel', 
                            font=('Segoe UI', 11), 
                            foreground=self.text_light,
                            background=self.card_bg)
        
        # Özel buton stilleri
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
        
        # Program butonları için stil
        self.style.configure('Program.TButton', 
                            font=('Segoe UI', 11),
                            padding=(10, 15))
        
        self.style.configure('ProgramSelected.TButton', 
                            font=('Segoe UI', 11, 'bold'),
                            padding=(10, 15))
        
        self.style.map('Program.TButton',
                      background=[('active', self.bg_color), 
                                 ('pressed', self.bg_color),
                                 ('!active', self.bg_color)],
                      foreground=[('active', self.primary_color), 
                                 ('pressed', self.primary_dark),
                                 ('!active', self.text_color)])
        
        self.style.map('ProgramSelected.TButton',
                      background=[('active', self.primary_color), 
                                 ('pressed', self.primary_dark),
                                 ('!active', self.primary_color)],
                      foreground=[('active', 'white'), 
                                 ('pressed', 'white'),
                                 ('!active', 'white')])
        
        # Programlar listesi
        self.programs = [
            {
                "name": "Aylık Fatura Doğrulama Programı",
                "description": "Excel dosyalarındaki faturaları karşılaştırarak doğrulama yapan program.",
                "icon": "📊",
                "module": "excel_karsilastir",
                "class": "ExcelKarsilastir",
                "path": "programs/fatura_dogrulama"
            },
            # Diğer programlar buraya eklenecek
        ]
        
        # Arayüz oluşturma
        self.create_widgets()
    
    def create_widgets(self):
        # Ana frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Üst panel (navbar)
        navbar = ttk.Frame(main_frame, style='TFrame', padding=(20, 15))
        navbar.pack(fill=tk.X)
        
        # Logo ve başlık
        logo_frame = ttk.Frame(navbar, style='TFrame')
        logo_frame.pack(side=tk.LEFT)
        
        # Logo (D harfi)
        logo_label = ttk.Label(logo_frame, 
                              text="D", 
                              font=('Segoe UI', 24, 'bold'),
                              foreground="white",
                              background=self.primary_color,
                              padding=(10, 0))
        logo_label.pack(side=tk.LEFT)
        
        # Başlık
        ttk.Label(logo_frame, 
                 text="Dokucuoglu Yazılım", 
                 style='Header.TLabel').pack(side=tk.LEFT, padx=(15, 0))
        
        # İçerik alanı - Scroll için canvas ekliyoruz
        content_container = ttk.Frame(main_frame, style='TFrame')
        content_container.pack(fill=tk.BOTH, expand=True)
        
        # Scroll için canvas oluştur
        canvas = tk.Canvas(content_container, bg=self.bg_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(content_container, orient="vertical", command=canvas.yview)
        
        # Scrollable frame
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        # Canvas içine frame yerleştir
        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Canvas genişliğini pencere genişliğine ayarla
        def _configure_canvas(event):
            canvas_width = event.width
            canvas.itemconfig(canvas_window, width=canvas_width)
        
        canvas.bind("<Configure>", _configure_canvas)
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Canvas ve scrollbar'ı yerleştir
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mouse wheel ile scroll
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # İçerik alanı
        content_frame = ttk.Frame(scrollable_frame, style='TFrame', padding=(20, 10, 20, 20))
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Tek sütunlu düzen yerine iki sütunlu grid kullanıyoruz
        # Sol panel için daha az ağırlık, sağ panel için daha fazla ağırlık
        content_frame.columnconfigure(0, weight=1)  # Sol panel için daha az ağırlık
        content_frame.columnconfigure(1, weight=4)  # Sağ panel için daha fazla ağırlık
        
        # Sol panel - Program listesi
        left_panel = ttk.Frame(content_frame, style='TFrame', padding=(0, 0, 15, 0))
        left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 20))
        
        # Programlar başlığı
        ttk.Label(left_panel, 
                 text="Programlar", 
                 style='Subheader.TLabel').pack(anchor="w", pady=(0, 15))
        
        # Program butonları için frame
        programs_frame = ttk.Frame(left_panel, style='TFrame')
        programs_frame.pack(fill=tk.BOTH)
        
        # Program butonları
        self.program_buttons = []
        for i, program in enumerate(self.programs):
            # Program butonu
            btn_style = 'ProgramSelected.TButton' if i == 0 else 'Program.TButton'
            program_btn = ttk.Button(
                programs_frame, 
                text=f"{program['icon']} {program['name']}", 
                style=btn_style,
                command=lambda p=program, idx=i: self.on_program_click(p, idx)
            )
            program_btn.pack(fill=tk.X, pady=5)
            self.program_buttons.append(program_btn)
        
        # Sağ panel - Program detayları
        self.right_panel = ttk.Frame(content_frame, style='TFrame')
        self.right_panel.grid(row=0, column=1, sticky="nsew")
        
        # Varsayılan olarak ilk programı göster
        if self.programs:
            self.show_program_details(self.programs[0])
        
        # Alt bilgi
        footer_frame = ttk.Frame(main_frame, style='TFrame', padding=(20, 10))
        footer_frame.pack(fill=tk.X)
        
        ttk.Label(footer_frame, 
                 text="© 2024 Dokucuoglu Yazılım", 
                 style='TLabel', 
                 foreground=self.text_light).pack(side=tk.LEFT)
        
        ttk.Label(footer_frame, 
                 text="v1.0.0", 
                 style='TLabel', 
                 foreground=self.text_light).pack(side=tk.RIGHT)
    
    def on_program_click(self, program, index):
        # Tüm butonları normal stile çevir
        for i, btn in enumerate(self.program_buttons):
            if i != index:
                btn.configure(style='Program.TButton')
        
        # Seçilen butonu vurgula
        self.program_buttons[index].configure(style='ProgramSelected.TButton')
        
        # Program detaylarını göster
        self.show_program_details(program)
    
    def show_program_details(self, program):
        # Sağ paneli temizle
        for widget in self.right_panel.winfo_children():
            widget.destroy()
        
        # Program detay kartı - Ortalanmış ve daha geniş
        detail_card = ttk.Frame(self.right_panel, style='Card.TFrame', padding=(30, 25))
        detail_card.pack(fill=tk.BOTH, expand=True, padx=20)
        
        # Kart için gölge efekti
        detail_card.configure(borderwidth=1, relief="solid")
        detail_card['borderwidth'] = 0
        
        # Program başlığı ve ikonu - Ortalanmış
        header_frame = ttk.Frame(detail_card, style='Card.TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 20), anchor="center")
        
        # İçeriği ortalamak için container
        center_container = ttk.Frame(header_frame, style='Card.TFrame')
        center_container.pack(anchor="center")
        
        # Program ikonu
        icon_label = ttk.Label(center_container, 
                              text=program["icon"], 
                              font=('Segoe UI', 48),  # Daha büyük ikon
                              style='Card.TLabel')
        icon_label.pack(side=tk.TOP, pady=(0, 10))
        
        # Program başlığı - Ortalanmış
        ttk.Label(center_container, 
                 text=program["name"], 
                 style='Title.TLabel').pack(anchor="center")
        
        # Program açıklaması - Ortalanmış
        description_frame = ttk.Frame(detail_card, style='Card.TFrame')
        description_frame.pack(fill=tk.X, pady=(0, 30))
        
        ttk.Label(description_frame, 
                 text=program["description"], 
                 style='Description.TLabel',
                 wraplength=600,  # Daha geniş açıklama
                 justify="center").pack(anchor="center", fill=tk.X)
        
        # Özellikler bölümü - Ortalanmış
        features_frame = ttk.Frame(detail_card, style='Card.TFrame')
        features_frame.pack(fill=tk.X, pady=(0, 30))
        
        ttk.Label(features_frame, 
                 text="Özellikler", 
                 font=('Segoe UI', 14, 'bold'),
                 foreground=self.text_color,
                 style='Card.TLabel').pack(anchor="center", pady=(0, 15))
        
        # Özellikler için container
        features_container = ttk.Frame(features_frame, style='Card.TFrame')
        features_container.pack(anchor="center")
        
        # Örnek özellikler
        features = [
            "Excel dosyalarını otomatik karşılaştırma",
            "Fark raporları oluşturma",
            "Hızlı ve doğru sonuçlar",
            "Kullanıcı dostu arayüz"
        ]
        
        for feature in features:
            feature_item = ttk.Frame(features_container, style='Card.TFrame')
            feature_item.pack(fill=tk.X, pady=5)
            
            # Özellik ikonu
            ttk.Label(feature_item, 
                     text="✓", 
                     foreground=self.primary_color,
                     font=('Segoe UI', 12, 'bold'),
                     style='Card.TLabel').pack(side=tk.LEFT)
            
            # Özellik metni
            ttk.Label(feature_item, 
                     text=feature, 
                     style='Card.TLabel',
                     padding=(5, 0, 0, 0)).pack(side=tk.LEFT, fill=tk.X)
        
        # Butonlar - Ortalanmış
        button_frame = ttk.Frame(detail_card, style='Card.TFrame')
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Butonları ortalamak için container
        button_container = ttk.Frame(button_frame, style='Card.TFrame')
        button_container.pack(anchor="center")
        
        # Çalıştır butonu
        run_button = ttk.Button(
            button_container, 
            text="Programı Çalıştır", 
            command=lambda: self.run_program(program),
            style='Primary.TButton'
        )
        run_button.pack(side=tk.LEFT)
        
        # Yardım butonu
        help_button = ttk.Button(
            button_container, 
            text="Yardım", 
            style='Secondary.TButton'
        )
        help_button.pack(side=tk.LEFT, padx=(10, 0))
    
    def run_program(self, program):
        try:
            # Program modülünün yolunu belirle
            program_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), program["path"])
            
            # Modül yolunu Python yoluna ekle
            if program_path not in sys.path:
                sys.path.append(program_path)
            
            # Modülü yükle
            module_name = program["module"]
            class_name = program["class"]
            
            # Modül zaten yüklenmişse, yeniden yükle
            if module_name in sys.modules:
                del sys.modules[module_name]
            
            # Modülü dinamik olarak içe aktar
            module = importlib.import_module(module_name)
            
            # Ana pencereyi gizle
            self.root.withdraw()
            
            # Yeni bir Tkinter penceresi oluştur
            program_window = tk.Toplevel(self.root)
            program_window.title(program["name"])
            program_window.geometry("900x650")
            program_window.protocol("WM_DELETE_WINDOW", lambda: self.on_program_close(program_window))
            
            # Program sınıfını başlat
            program_class = getattr(module, class_name)
            program_instance = program_class(program_window)
            
            # Pencere kapandığında ana pencereyi göster
            program_window.mainloop()
            
        except Exception as e:
            messagebox.showerror("Hata", f"Program çalıştırılırken bir hata oluştu:\n{str(e)}")
            self.root.deiconify()  # Ana pencereyi tekrar göster
    
    def on_program_close(self, program_window):
        program_window.destroy()
        self.root.deiconify()  # Ana pencereyi tekrar göster

def main():
    root = tk.Tk()
    app = DokucuogluYazilim(root)
    
    # Uygulama simgesi
    try:
        if platform.system() == "Windows":
            root.iconbitmap("icon.ico")
    except:
        pass
    
    # Pencere arkaplan rengi
    root.configure(bg="#f8fafc")
    
    root.mainloop()

if __name__ == "__main__":
    main() 
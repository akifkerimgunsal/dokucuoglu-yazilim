# Fatura Doğrulama Programı paketi
# Bu dosya, Python'un bu dizini bir paket olarak tanımasını sağlar

# Gerekli modülleri içe aktar
import sys
import os

# Mevcut dizini Python yoluna ekle
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.append(current_dir) 
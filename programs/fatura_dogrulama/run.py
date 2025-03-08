#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Aylık Fatura Doğrulama Programı
Bu dosya, programı doğrudan çalıştırmak için kullanılır.
"""

import tkinter as tk
from excel_karsilastir import ExcelKarsilastir

def main():
    root = tk.Tk()
    app = ExcelKarsilastir(root)
    root.mainloop()

if __name__ == "__main__":
    main() 
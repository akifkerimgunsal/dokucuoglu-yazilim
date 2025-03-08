# Dokucuoglu Yazılım - Program Merkezi

![Dokucuoglu Yazılım](https://img.shields.io/badge/Dokucuoglu-Yazılım-blue)
![Versiyon](https://img.shields.io/badge/Versiyon-1.0.0-green)
![Python](https://img.shields.io/badge/Python-3.8+-yellow)

Dokucuoglu Yazılım Program Merkezi, şirket içi kullanım için geliştirilmiş çeşitli iş süreçlerini otomatikleştiren ve kolaylaştıran programları tek bir arayüz altında toplayan bir uygulamadır.

## 📋 İçindekiler

- [Genel Bakış](#genel-bakış)
- [Özellikler](#özellikler)
- [Kurulum](#kurulum)
- [Kullanım](#kullanım)
- [Mevcut Programlar](#mevcut-programlar)


## 🔍 Genel Bakış

Dokucuoglu Yazılım Program Merkezi, şirket içinde kullanılan çeşitli programları tek bir merkezi arayüzden erişilebilir hale getiren bir uygulamadır. Modern ve kullanıcı dostu arayüzü sayesinde, kullanıcılar ihtiyaç duydukları programlara kolayca erişebilir ve işlemlerini hızlıca gerçekleştirebilirler.

Program Merkezi, modüler bir yapıya sahiptir ve yeni programlar kolayca eklenebilir. Her program, kendi dizininde bağımsız olarak çalışabilir ve Program Merkezi üzerinden erişilebilir.

## ✨ Özellikler

- **Modern ve Kullanıcı Dostu Arayüz**: Sezgisel ve estetik bir kullanıcı deneyimi sunar
- **Modüler Yapı**: Yeni programlar kolayca eklenebilir ve mevcut programlar güncellenebilir
- **Merkezi Erişim**: Tüm programlara tek bir arayüzden erişim sağlar
- **Tam Ekran Desteği**: Geniş ekranlarda optimum kullanım için tam ekran modu
- **Duyarlı Tasarım**: Farklı ekran boyutlarına uyum sağlayan esnek arayüz

## 💻 Kurulum

### Gereksinimler

- Python 3.8 veya üzeri
- Gerekli Python paketleri (requirements.txt dosyasında listelenmiştir)

### Kurulum Adımları

1. Depoyu klonlayın veya indirin:
   ```
   git clone https://github.com/dokucuoglu/dokucuoglu-yazilim.git
   ```

2. Proje dizinine gidin:
   ```
   cd dokucuoglu-yazilim
   ```

3. Gerekli paketleri yükleyin:
   ```
   pip install -r requirements.txt
   ```

## 🚀 Kullanım

Program Merkezi'ni başlatmak için:

1. Windows'ta `calistir.vbs` dosyasına çift tıklayın
   
   veya
   
2. Komut satırından Python ile çalıştırın:
   ```
   python main.py
   ```

Program başladığında, sol panelde mevcut programların listesini göreceksiniz. Bir programa tıkladığınızda, sağ panelde o programın detayları ve çalıştırma seçenekleri görüntülenecektir.

## Mevcut Programlar

### 1. Aylık Fatura Doğrulama Programı

Excel dosyalarındaki faturaları karşılaştırarak doğrulama yapmanızı sağlayan bir programdır.

**Özellikler:**
- Excel dosyalarını karşılaştırma
- Sayısal değerlerdeki farkları yüzde olarak gösterme
- Para birimi uyuşmazlıklarını tespit etme
- Detaylı Excel raporu oluşturma

**Kullanım:**
1. "Gelen Fatura" ve "İşlenmiş Fatura" dosyalarını seçin
2. Her iki dosya için karşılaştırmak istediğiniz sütunları seçin (en az 4 sütun)
3. "Rapor Oluştur" butonuna tıklayarak karşılaştırma raporunu oluşturun


© 2024 Dokucuoglu Yazılım. Tüm hakları saklıdır.

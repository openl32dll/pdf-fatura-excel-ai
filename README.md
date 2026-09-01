# PDF Fatura → Excel + AI Destekli Kategorize Aracı

PDF faturaları okur, kalemleri (açıklama / miktar / birim fiyat / tutar) otomatik
olarak tablo halinde çıkarır, her kalemi bir muhasebe kategorisine ayırır ve
biçimlendirilmiş, iki sayfalı (Detay + Özet) bir Excel raporu üretir.

Bu proje bir **Fiverr/Upwork gig'i için portföy/demo** amacıyla hazırlanmıştır:
"PDF faturadan Excel'e otomatik veri aktarma" ve "AI destekli Excel veri
temizleme/kategorize etme" hizmetlerinin uçtan uca çalışan bir kanıtıdır.

**Çıktı dili** (`--lang en` / `--lang tr`) faturanın hangi dilde olduğundan
bağımsızdır — varsayılan İngilizce (uluslararası müşteriler için), Türkçe
müşteriler için `--lang tr` ile kolonlar ve kategoriler Türkçeye döner.

## Özellikler

- 📄 PDF'deki fatura tablosunu otomatik tanır ve satırlara ayırır
- 🏷️ Her kalemi kategoriye ayırır — **Ofis Malzemesi, Yazılım/Lisans, Ulaşım,
  Yemek/İkram, Danışmanlık, Diğer**
- 🤖 `ANTHROPIC_API_KEY` tanımlıysa kategorizasyon **Claude AI** ile yapılır
  (karışık/çeşitli fatura kalemlerinde çok daha isabetli sonuç verir)
- 🔁 API anahtarı yoksa otomatik olarak anahtar-kelime tabanlı moda düşer —
  yani araç **AI olmadan da uçtan uca çalışır**
- 📊 Çıktı: "Fatura Detay" (satır satır döküm) + "Özet" (kategori bazlı
  toplamlar, genel toplam) sayfalarından oluşan biçimlendirilmiş bir `.xlsx`

## Kurulum

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Kullanım

```bash
# İki demo fatura üretir: sample_data/invoice_sample_en.pdf (EN) ve
# sample_data/fatura_ornek.pdf (TR)
python generate_sample.py

# İngilizce çıktı (varsayılan) — AI varsa otomatik kullanılır
export ANTHROPIC_API_KEY=sk-...   # opsiyonel
python extract.py sample_data/invoice_sample_en.pdf -o output/report.xlsx

# Türkçe çıktı için:
python extract.py sample_data/fatura_ornek.pdf -o output/rapor.xlsx --lang tr

# AI'yı zorla kapatmak için:
python extract.py sample_data/invoice_sample_en.pdf -o output/report.xlsx --no-ai
```

Çıktı belirttiğin `-o` yoluna oluşur: "Detay" sayfasında her kalem + atanmış
kategorisi, "Özet" sayfasında fatura bilgileri ve kategori bazlı toplamlar
yer alır — sayfa/kolon isimleri seçtiğin `--lang`'a göre değişir.

## Nasıl çalışır?

1. **`pdfplumber`** ile PDF'deki tablo yapısı (kenarlıklar/gridler) tespit
   edilip satır satır çıkarılır — düz metin regex'ine güvenmek yerine PDF'in
   gerçek tablo yapısını okur, bu da farklı fatura formatlarına karşı çok
   daha sağlam çalışır.
2. Fatura başlığından (No, Tarih, Satıcı) bilgiler regex ile çekilir.
3. Her kalem, ya Claude API ile (anlam bazlı, "Bulut Depolama Aboneliği"
   → Yazılım/Lisans gibi belirsiz durumları da doğru ayırt eder) ya da
   basit anahtar kelime eşlemesiyle kategorize edilir.
4. `openpyxl` ile biçimlendirilmiş, iki sayfalı bir Excel raporu üretilir.

## Sınırlamalar / sonraki adımlar

- Şu an tablo yapısı belirgin (kenarlıklı) PDF'ler için optimize edilmiştir;
  taranmış (resim) faturalar için OCR (ör. `pytesseract`) eklenmesi gerekir.
- Farklı fatura düzenlerine (çok satırlı açıklamalar, farklı sütun sırası)
  uyum için sütun eşleme mantığı genişletilebilir.
- Toplu işleme (bir klasördeki tüm PDF'leri tek seferde işleyip tek Excel'de
  birleştirme) kolayca eklenebilir.

## Lisans

MIT

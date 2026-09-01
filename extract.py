#!/usr/bin/env python3
"""
PDF Fatura -> Excel + AI Destekli Kategorize Aracı
====================================================

Bir PDF faturayı okur, kalem satırlarını (açıklama/miktar/birim fiyat/tutar)
tablo olarak çıkarır, her kalemi bir muhasebe kategorisine ayırır ve
biçimlendirilmiş bir Excel raporu (detay + özet sayfası) üretir.

İki mod:
  - AI modu:  ANTHROPIC_API_KEY ortam değişkeni ayarlıysa, kategorizasyon
              Claude ile yapılır (açıklaması belirsiz/çeşitli kalemlerde
              çok daha isabetli sonuç verir).
  - Fallback: API anahtarı yoksa basit anahtar kelime eşlemesiyle
              kategorize eder — API key olmadan da uçtan uca çalışır.

Kullanım:
    python extract.py sample_data/fatura_ornek.pdf -o output/rapor.xlsx
    python extract.py fatura.pdf -o rapor.xlsx --no-ai   # AI'yı zorla kapat
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

CATEGORIES = [
    "Ofis Malzemesi",
    "Yazılım/Lisans",
    "Ulaşım",
    "Yemek/İkram",
    "Danışmanlık",
    "Diğer",
]

KEYWORD_MAP = {
    "Yazılım/Lisans": ["yazılım", "lisans", "bulut", "abonelik", "software", "subscription"],
    "Ulaşım": ["taksi", "otobüs", "uçak", "yakıt", "benzin", "ulaşım", "transfer"],
    "Yemek/İkram": ["yemek", "kahve", "ikram", "catering", "restoran"],
    "Danışmanlık": ["danışman", "consulting", "hizmet bedeli"],
    "Ofis Malzemesi": ["kağıt", "kalem", "mouse", "klavye", "ofis", "toner", "dosya"],
}


# --------------------------------------------------------------------------
# PDF'den ham veri çıkarımı
# --------------------------------------------------------------------------

def extract_text(pdf_path: Path) -> str:
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


def extract_line_items(pdf_path: Path) -> list[dict]:
    items = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            for table in page.extract_tables():
                if not table:
                    continue
                header = [c.strip().lower() if c else "" for c in table[0]]
                if not any("açıklama" in h or "aciklama" in h for h in header):
                    continue
                for row in table[1:]:
                    if not row or len(row) < 4 or not row[0]:
                        continue
                    desc, qty, unit_price, total = row[:4]
                    items.append(
                        {
                            "aciklama": desc.strip(),
                            "miktar": _to_number(qty),
                            "birim_fiyat": _to_number(unit_price),
                            "tutar": _to_number(total),
                        }
                    )
    return items


def extract_header_fields(text: str) -> dict:
    def find(pattern, default=""):
        m = re.search(pattern, text, flags=re.IGNORECASE)
        return m.group(1).strip() if m else default

    return {
        "fatura_no": find(r"Fatura No:?\s*(.+)"),
        "tarih": find(r"Tarih:?\s*([\d.\/\-]+)"),
        "satici": find(r"Sat[ıi]c[ıi]:?\s*(.+)"),
    }


def _to_number(raw: str) -> float:
    if raw is None:
        return 0.0
    cleaned = raw.strip().replace(".", "").replace(",", ".") if "," in raw else raw.strip()
    try:
        return float(cleaned)
    except ValueError:
        try:
            return float(raw.strip())
        except ValueError:
            return 0.0


# --------------------------------------------------------------------------
# Kategorizasyon
# --------------------------------------------------------------------------

def categorize_with_keywords(items: list[dict]) -> list[dict]:
    for item in items:
        desc_lower = item["aciklama"].lower()
        category = "Diğer"
        for cat, keywords in KEYWORD_MAP.items():
            if any(kw in desc_lower for kw in keywords):
                category = cat
                break
        item["kategori"] = category
    return items


def categorize_with_ai(items: list[dict]) -> list[dict]:
    import anthropic

    client = anthropic.Anthropic()
    descriptions = [item["aciklama"] for item in items]

    prompt = f"""Aşağıdaki fatura kalemlerinin her birini şu kategorilerden
birine ata: {", ".join(CATEGORIES)}.

Kalemler (sırasıyla):
{json.dumps(descriptions, ensure_ascii=False, indent=2)}

SADECE şu formatta bir JSON listesi döndür, başka hiçbir açıklama ekleme:
["Kategori1", "Kategori2", ...]
Liste uzunluğu tam olarak kalem sayısı kadar olmalı ve sırayı korumalı."""

    resp = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=1000,
        messages=[{"role": "user", "content": prompt}],
    )
    raw = resp.content[0].text.strip()
    raw = re.sub(r"^```(json)?|```$", "", raw, flags=re.MULTILINE).strip()
    categories = json.loads(raw)

    for item, cat in zip(items, categories):
        item["kategori"] = cat if cat in CATEGORIES else "Diğer"
    return items


# --------------------------------------------------------------------------
# Excel raporu
# --------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def build_excel(header: dict, items: list[dict], out_path: Path) -> None:
    wb = Workbook()

    # --- Detay sayfası ---
    ws = wb.active
    ws.title = "Fatura Detay"
    cols = ["Açıklama", "Miktar", "Birim Fiyat", "Tutar", "Kategori"]
    ws.append(cols)
    for c in range(1, len(cols) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")

    for item in items:
        ws.append(
            [item["aciklama"], item["miktar"], item["birim_fiyat"], item["tutar"], item["kategori"]]
        )

    for c in range(1, len(cols) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 24
    for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
        for cell in row:
            cell.number_format = "#,##0.00 TL"

    # --- Özet sayfası ---
    ws2 = wb.create_sheet("Özet")
    ws2.append(["Fatura No", header.get("fatura_no", "")])
    ws2.append(["Tarih", header.get("tarih", "")])
    ws2.append(["Satıcı", header.get("satici", "")])
    ws2.append([])
    ws2.append(["Kategori", "Toplam Tutar"])
    for c in range(5, 7):
        cell = ws2.cell(row=5, column=c - 4)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL

    totals: dict[str, float] = {}
    for item in items:
        totals[item["kategori"]] = totals.get(item["kategori"], 0.0) + item["tutar"]

    for cat, total in sorted(totals.items(), key=lambda kv: -kv[1]):
        ws2.append([cat, total])
    ws2.append([])
    ws2.append(["GENEL TOPLAM", sum(item["tutar"] for item in items)])
    ws2["A" + str(ws2.max_row)].font = Font(bold=True)
    ws2["B" + str(ws2.max_row)].font = Font(bold=True)

    for row in ws2.iter_rows(min_row=6, min_col=2, max_col=2):
        for cell in row:
            cell.number_format = "#,##0.00 TL"
    ws2.column_dimensions["A"].width = 24
    ws2.column_dimensions["B"].width = 18

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="PDF faturayı Excel raporuna dönüştürür.")
    parser.add_argument("pdf", type=Path, help="Girdi PDF fatura dosyası")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/rapor.xlsx"))
    parser.add_argument("--no-ai", action="store_true", help="AI kategorizasyonu kullanma")
    args = parser.parse_args()

    if not args.pdf.exists():
        sys.exit(f"Hata: {args.pdf} bulunamadı")

    text = extract_text(args.pdf)
    header = extract_header_fields(text)
    items = extract_line_items(args.pdf)

    if not items:
        sys.exit("Hata: PDF içinde tanınabilir bir fatura tablosu bulunamadı.")

    use_ai = not args.no_ai
    if use_ai:
        try:
            items = categorize_with_ai(items)
            mode = "AI (Claude)"
        except Exception as e:  # API anahtarı yok / hata -> fallback
            print(f"[uyarı] AI kategorizasyon kullanılamadı ({e}); anahtar kelime moduna geçiliyor.")
            items = categorize_with_keywords(items)
            mode = "Anahtar kelime (fallback)"
    else:
        items = categorize_with_keywords(items)
        mode = "Anahtar kelime (--no-ai)"

    build_excel(header, items, args.output)

    print(f"✅ Rapor oluşturuldu: {args.output}")
    print(f"   Kategorizasyon modu: {mode}")
    print(f"   {len(items)} kalem işlendi, genel toplam: {sum(i['tutar'] for i in items):,.2f} TL")


if __name__ == "__main__":
    main()

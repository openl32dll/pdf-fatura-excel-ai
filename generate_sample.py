#!/usr/bin/env python3
"""
Generates demo/test invoice PDFs for extract.py — not real customer data,
just a sample to demonstrate how the tool works.

Produces two files:
  sample_data/invoice_sample_en.pdf  (English, US-style numbers: 1,200.00)
  sample_data/fatura_ornek.pdf       (Turkish, EU-style numbers: 1.200,00)

Usage:
    python generate_sample.py
"""
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet

# Standard Helvetica doesn't render Turkish characters (ı, ş, ğ, ç, ö, ü)
# correctly -> register a Unicode TrueType font. Bundled in the repo
# (fonts/) so this works identically on Windows/Mac/Linux with no system
# font dependency.
FONT_DIR = Path(__file__).parent / "fonts"
pdfmetrics.registerFont(TTFont("DejaVuSans", FONT_DIR / "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", FONT_DIR / "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFontFamily("DejaVuSans", normal="DejaVuSans", bold="DejaVuSans-Bold")

OUT_DIR = Path(__file__).parent / "sample_data"
OUT_DIR.parent.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

styles = getSampleStyleSheet()
styles["Normal"].fontName = "DejaVuSans"


def build_invoice(out_path: Path, header_text: str, table_data: list[list[str]]) -> None:
    doc = SimpleDocTemplate(str(out_path), pagesize=A4, topMargin=2 * cm, bottomMargin=2 * cm)
    table = Table(table_data, colWidths=[8 * cm, 2.5 * cm, 3 * cm, 3 * cm])
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#dddddd")),
                ("FONTNAME", (0, 0), (-1, -1), "DejaVuSans"),
                ("FONTNAME", (0, 0), (-1, 0), "DejaVuSans-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story = [Paragraph(header_text, styles["Normal"]), Spacer(1, 1 * cm), table]
    doc.build(story)
    print(f"Created: {out_path}")


# --- English sample (US-style numbers: comma thousands, dot decimal) ---
en_header = """
<b>INVOICE</b><br/>
Invoice No: INV-2026-0451<br/>
Date: 08/15/2026<br/>
Vendor: Example Trading Inc.<br/>
Bill To: Demo Client LLC
"""
en_table = [
    ["Description", "Qty", "Unit Price", "Total"],
    ["Wireless Mouse", "3", "250.00", "750.00"],
    ["Software License (Annual)", "1", "1,200.00", "1,200.00"],
    ["Office Paper (A4, 5 packs)", "2", "180.00", "360.00"],
    ["Taxi Fare", "1", "95.00", "95.00"],
    ["Consulting Service (2 hrs)", "2", "500.00", "1,000.00"],
    ["Office Coffee/Catering", "1", "220.00", "220.00"],
    ["Cloud Storage Subscription", "1", "340.00", "340.00"],
]
build_invoice(OUT_DIR / "invoice_sample_en.pdf", en_header, en_table)

# --- Turkish sample (EU-style numbers: dot thousands, comma decimal) ---
tr_header = """
<b>FATURA</b><br/>
Fatura No: FAT-2026-0451<br/>
Tarih: 15.08.2026<br/>
Satıcı: Örnek Ticaret A.Ş.<br/>
Alıcı: Demo Müşteri Ltd. Şti.
"""
tr_table = [
    ["Açıklama", "Miktar", "Birim Fiyat", "Tutar"],
    ["Kablosuz Mouse", "3", "250,00", "750,00"],
    ["Yazılım Lisansı (Yıllık)", "1", "1.200,00", "1.200,00"],
    ["Ofis Kağıdı (A4, 5 Paket)", "2", "180,00", "360,00"],
    ["Taksi Ücreti", "1", "95,00", "95,00"],
    ["Danışmanlık Hizmeti (2 saat)", "2", "500,00", "1.000,00"],
    ["Ofis Kahve/İkram", "1", "220,00", "220,00"],
    ["Bulut Depolama Aboneliği", "1", "340,00", "340,00"],
]
build_invoice(OUT_DIR / "fatura_ornek.pdf", tr_header, tr_table)

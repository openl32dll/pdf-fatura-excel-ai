#!/usr/bin/env python3
"""
Demo/test amaçlı örnek bir fatura PDF'i üretir.
Gerçek müşteri PDF'i değildir — sadece extract.py'nin nasıl çalıştığını
göstermek için kullanılır.

Kullanım:
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

# Standart Helvetica fontu Türkçe karakterleri (ı, ş, ğ, ç, ö, ü) doğru
# encode etmiyor -> Unicode destekli bir TrueType font kaydediyoruz.
FONT_DIR = Path("/usr/share/fonts/truetype/dejavu")
pdfmetrics.registerFont(TTFont("DejaVuSans", FONT_DIR / "DejaVuSans.ttf"))
pdfmetrics.registerFont(TTFont("DejaVuSans-Bold", FONT_DIR / "DejaVuSans-Bold.ttf"))
pdfmetrics.registerFontFamily("DejaVuSans", normal="DejaVuSans", bold="DejaVuSans-Bold")

OUT = Path(__file__).parent / "sample_data" / "fatura_ornek.pdf"
OUT.parent.mkdir(exist_ok=True)

styles = getSampleStyleSheet()
styles["Normal"].fontName = "DejaVuSans"
doc = SimpleDocTemplate(str(OUT), pagesize=A4, topMargin=2 * cm, bottomMargin=2 * cm)

header_text = """
<b>FATURA</b><br/>
Fatura No: FAT-2026-0451<br/>
Tarih: 15.08.2026<br/>
Satıcı: Örnek Ticaret A.Ş.<br/>
Alıcı: Demo Müşteri Ltd. Şti.
"""

table_data = [
    ["Açıklama", "Miktar", "Birim Fiyat", "Tutar"],
    ["Kablosuz Mouse", "3", "250.00", "750.00"],
    ["Yazılım Lisansı (Yıllık)", "1", "1200.00", "1200.00"],
    ["Ofis Kağıdı (A4, 5 Paket)", "2", "180.00", "360.00"],
    ["Taksi Ücreti", "1", "95.00", "95.00"],
    ["Danışmanlık Hizmeti (2 saat)", "2", "500.00", "1000.00"],
    ["Ofis Kahve/İkram", "1", "220.00", "220.00"],
    ["Bulut Depolama Aboneliği", "1", "340.00", "340.00"],
]

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
print(f"Örnek fatura oluşturuldu: {OUT}")

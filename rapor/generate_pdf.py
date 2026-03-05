#!/usr/bin/env python3
"""Plastik Enjeksiyon Proses Kontrol Raporu - PDF Oluşturucu (ReportLab)"""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT

DARK_BLUE = HexColor('#29417A')
LIGHT_BLUE = HexColor('#DCE6F1')
ALT_ROW = HexColor('#F0F5FA')
WHITE = white
BLACK = black

OUTPUT = '/home/user/kontrol/rapor/PROSES_KONTROL_RAPORU.pdf'


def make_style(name, font='Helvetica', size=7, alignment=TA_LEFT, bold=False, color=BLACK):
    styles = getSampleStyleSheet()
    s = styles['Normal'].clone(name)
    s.fontName = f'{font}-Bold' if bold else font
    s.fontSize = size
    s.leading = size + 2
    s.alignment = alignment
    s.textColor = color
    return s

TITLE_S = make_style('t', size=12, bold=True, alignment=TA_CENTER)
SUB_S = make_style('sub', size=7, alignment=TA_CENTER)
SECT_S = make_style('sect', size=9, bold=True, color=WHITE)
SUBSECT_S = make_style('subsect', size=8, bold=True)
CELL_S = make_style('cell', size=7, alignment=TA_CENTER)
CELL_L = make_style('celll', size=7, alignment=TA_LEFT)
CELL_B = make_style('cellb', size=7, alignment=TA_CENTER, bold=True)
NOTE_S = make_style('note', size=6, alignment=TA_LEFT)
FOOTER_S = make_style('footer', size=6, alignment=TA_CENTER)


def p(text, style=CELL_S):
    return Paragraph(str(text), style)


def section_banner(text):
    data = [[p(text, SECT_S)]]
    t = Table(data, colWidths=[270*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), DARK_BLUE),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
    ]))
    return t


def sub_banner(text):
    data = [[p(text, SUBSECT_S)]]
    t = Table(data, colWidths=[270*mm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), LIGHT_BLUE),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
    ]))
    return t


def styled_table(headers, rows, col_widths=None):
    if col_widths is None:
        w = 270*mm / len(headers)
        col_widths = [w] * len(headers)

    data = [[p(h, CELL_B) for h in headers]]
    for row in rows:
        data.append([p(c, CELL_S) for c in row])

    style_cmds = [
        ('BACKGROUND', (0,0), (-1,0), DARK_BLUE),
        ('TEXTCOLOR', (0,0), (-1,0), WHITE),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#CCCCCC')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]
    for i in range(1, len(data)):
        if i % 2 == 0:
            style_cmds.append(('BACKGROUND', (0,i), (-1,i), ALT_ROW))

    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle(style_cmds))
    return t


def info_pair_table(left_rows, right_rows):
    """Two-column info table side by side."""
    left_data = []
    for label, val in left_rows:
        left_data.append([p(label, CELL_B), p(val, CELL_L)])
    right_data = []
    for label, val in right_rows:
        right_data.append([p(label, CELL_B), p(val, CELL_L)])

    lt = Table(left_data, colWidths=[40*mm, 60*mm])
    lt.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#CCCCCC')),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
    ]))

    rt = Table(right_data, colWidths=[42*mm, 60*mm])
    rt.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, HexColor('#CCCCCC')),
        ('TOPPADDING', (0,0), (-1,-1), 1),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1),
    ]))

    outer = Table([[lt, rt]], colWidths=[105*mm, 107*mm])
    outer.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    return outer


def build_report():
    doc = SimpleDocTemplate(
        OUTPUT,
        pagesize=landscape(A4),
        leftMargin=10*mm, rightMargin=10*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )
    story = []
    sp = Spacer(1, 3*mm)
    sp2 = Spacer(1, 5*mm)

    # Title
    story.append(p('PLASTIK ENJEKSIYON PROSES KONTROL RAPORU', TITLE_S))
    story.append(p('Dokuman No: KKR-ENJ-001  |  Rev: 01  |  Ref: ISO 9001:2015 Madde 8.5.1 / ISO 2859-1 (AQL)', SUB_S))
    story.append(sp2)

    # ── BOLUM A ──
    story.append(section_banner('BOLUM A: GENEL BILGILER'))
    story.append(sp)
    story.append(info_pair_table(
        [
            ('Musteri', 'EBERLE'),
            ('Urun Adi', 'KARE KCK K.M'),
            ('Siparis Adedi', '100.000 Adet'),
            ('Makine', 'BORCHE (Proses 6)'),
            ('Kalip No', 'EBERLE.KARE.KCK.K.M'),
            ('Hammadde Tipi', ''),
        ],
        [
            ('Urun Kodu', ''),
            ('Siparis No', ''),
            ('Parti / Lot No', ''),
            ('Hammadde Parti No', ''),
            ('Rapor Tarihi', '.../.../202...'),
            ('Kontrol Sikligi', 'Her 50 Baski'),
        ]
    ))
    story.append(sp2)

    # ── BOLUM B ──
    story.append(section_banner('BOLUM B: MAKINE PARAMETRELERI'))
    story.append(sp)

    # B.1
    story.append(sub_banner('B.1 - Enjeksiyon Parametreleri (4 Asama)'))
    story.append(sp)
    story.append(styled_table(
        ['Asama', 'Basinc (bar)', 'Hiz (%)', 'Pozisyon (mm)'],
        [
            ['1', '161.0', '80.0', '11.0'],
            ['2', '155.0', '45.0', '10.0'],
            ['3', '140.0', '40.0', '9.0'],
            ['4', '160.0', '35.0', '-'],
        ],
        col_widths=[25*mm, 60*mm, 60*mm, 60*mm]
    ))
    story.append(sp)

    # B.2
    story.append(sub_banner('B.2 - Tutma (Holding) Parametreleri (3 Asama)'))
    story.append(sp)
    story.append(styled_table(
        ['Asama', 'Basinc (bar)', 'Hiz (%)', 'Sure (s)'],
        [
            ['1', '60.0', '5.0', '1.00'],
            ['2', '0.0', '0.0', '0.00'],
            ['3', '0.0', '0.0', '0.00'],
        ],
        col_widths=[25*mm, 65*mm, 65*mm, 65*mm]
    ))
    story.append(sp)

    # B.3
    story.append(sub_banner('B.3 - Plastiklestirme (Mal Alma) Parametreleri (3 Asama)'))
    story.append(sp)
    story.append(styled_table(
        ['Asama', 'Geri Basinc (bar)', 'Basinc (bar)', 'Hiz (%)', 'Pozisyon (mm)'],
        [
            ['1', '5.0', '90', '70', '10.0'],
            ['2', '8.0', '90', '70', '15.0'],
            ['3', '10.0', '90', '70', '28.0'],
        ],
        col_widths=[22*mm, 55*mm, 55*mm, 55*mm, 55*mm]
    ))
    story.append(sp)

    # B.4
    story.append(sub_banner('B.4 - Sicaklik Profili (Isitma Bolgeleri)'))
    story.append(sp)
    story.append(styled_table(
        ['Bolge 1 (Nozul)', 'Bolge 2', 'Bolge 3', 'Bolge 4', 'Bolge 5'],
        [['275 C', '270 C', '270 C', '271 C', '270 C']],
        col_widths=[50*mm, 50*mm, 50*mm, 50*mm, 50*mm]
    ))
    story.append(sp)
    story.append(styled_table(
        ['Parametre', 'Deger'],
        [
            ['Bekleme Isisi', '100 C'],
            ['Yag Sicakligi', '21 C'],
        ],
        col_widths=[130*mm, 130*mm]
    ))
    story.append(sp2)

    # B.5
    story.append(sub_banner('B.5 - Ozet Proses Parametreleri'))
    story.append(sp)
    cw9 = [8*mm, 50*mm, 16*mm, 22*mm, 22*mm, 22*mm, 28*mm, 28*mm, 28*mm, 22*mm]
    story.append(styled_table(
        ['No', 'Parametre', 'Birim', 'Alt Limit', 'Hedef', 'Ust Limit', 'Gercek 1', 'Gercek 2', 'Gercek 3', 'Durum'],
        [
            ['1', 'Kalip Sicakligi', 'C', '-', '-', '-', '-', '-', '-', '-'],
            ['2', 'Eriyik Sicakligi', 'C', '270', '272', '275', '', '', '', ''],
            ['3', 'Enj. Basinc (As.1)', 'bar', '155', '161', '165', '161.0', '161.0', '', ''],
            ['4', 'Enj. Hiz (As.1)', '%', '75', '80', '85', '80.0', '80.0', '', ''],
            ['5', 'Utuleme Basinci', 'bar', '55', '60', '65', '60.0', '60.0', '', ''],
            ['6', 'Utuleme Suresi', 's', '0.8', '1.0', '1.2', '1.00', '1.00', '', ''],
            ['7', 'Sogutma Suresi', 's', '23', '25', '27', '25.0', '25.0', '', ''],
            ['8', 'Cevrim Suresi', 's', '25', '27.6', '35', '27.6', '49.7*', '', ''],
            ['9', 'Sikma Kuvveti', 'ton', '-', '-', '-', '-', '-', '-', '-'],
        ],
        col_widths=cw9
    ))
    story.append(Spacer(1, 1*mm))
    story.append(p('* Cevrim suresi 49.7s: Olagandisi yuksek - arastirilmali (durus, kalip acma gecikmesi vb.)', NOTE_S))
    story.append(sp)

    # B.6
    story.append(sub_banner('B.6 - Ek Kesme / Zamanlama Parametreleri'))
    story.append(sp)
    story.append(styled_table(
        ['Parametre', 'Deger'],
        [
            ['Enjek Basinc (Gercek)', '159.4 - 159.5 bar'],
            ['Enjek Zaman', '4.00 s'],
            ['Yastiklama (Cushion)', '6.36 - 6.45 mm'],
            ['Anlik Tutma Zamani', '1.0 s'],
            ['Cut-off Pozisyonu', '21.6 mm'],
            ['Vida Pozisyonu', '31.26 - 31.30 mm'],
        ],
        col_widths=[130*mm, 130*mm]
    ))

    # ── PAGE 2 ──
    story.append(PageBreak())

    # ── BOLUM C ──
    story.append(section_banner('BOLUM C: GORSEL VE BOYUTSAL KALITE KONTROL (ISO 2859-1 AQL)'))
    story.append(sp)
    cw12 = [10*mm, 18*mm, 22*mm, 18*mm, 20*mm, 20*mm, 22*mm, 22*mm, 20*mm, 20*mm, 20*mm, 38*mm]
    rows_c = [[str(i), '', '', '', '', '', '', '', '', '', '', ''] for i in range(1, 13)]
    story.append(styled_table(
        ['No', 'Saat', 'Uretim', 'Numune', 'Cokuntu', 'Capak', 'Cizik', 'K.Yuzey', 'Renk', 'Boyut', 'Sonuc', 'Imza'],
        rows_c,
        col_widths=cw12
    ))
    story.append(sp2)

    # ── BOLUM D ──
    story.append(section_banner('BOLUM D: URETIM OZETI'))
    story.append(sp)
    story.append(styled_table(
        ['Parametre', 'Deger'],
        [
            ['Toplam Uretim (Adet)', ''],
            ['Sevk Edilen (Adet)', ''],
            ['Hatali / Fire (Adet)', ''],
            ['Fire Orani (%)', '0,00%'],
            ['Kabul Orani (%)', '0,00%'],
            ['AQL Seviyesi', '1.0 (Genel Kontrol Seviyesi II)'],
        ],
        col_widths=[130*mm, 130*mm]
    ))
    story.append(sp2)

    # ── BOLUM E ──
    story.append(section_banner('BOLUM E: HATA TURLERI ANALIZI'))
    story.append(sp)
    cw_e = [10*mm, 55*mm, 25*mm, 25*mm, 40*mm, 95*mm]
    story.append(styled_table(
        ['No', 'Hata Turu', 'Adet', 'Oran (%)', 'Sinif', 'Duzeltici Faaliyet'],
        [
            ['1', 'Cokuntu Izi (Sink Mark)', '', '0,0%', '', ''],
            ['2', 'Capak (Flash)', '', '0,0%', '', ''],
            ['3', 'Kisa Baski (Short Shot)', '', '0,0%', '', ''],
            ['4', 'Yanik Izi (Burn Mark)', '', '0,0%', '', ''],
            ['5', 'Egilme / Carpilma (Warpage)', '', '0,0%', '', ''],
            ['6', 'Cizik / Leke', '', '0,0%', '', ''],
            ['7', 'Renk Uyumsuzlugu', '', '0,0%', '', ''],
            ['8', 'Diger', '', '0,0%', '', ''],
            ['', 'TOPLAM HATA', '0', '0,0%', '', ''],
        ],
        col_widths=cw_e
    ))
    story.append(sp2)

    # ── BOLUM F ──
    story.append(section_banner('BOLUM F: KARAR VE ONAY'))
    story.append(sp)
    story.append(styled_table(
        ['PARTI KARARI'],
        [['[ ] KABUL          [ ] SARTLI KABUL          [ ] RED']],
        col_widths=[270*mm]
    ))
    story.append(styled_table(
        ['Aciklama / Sartli Kabul Nedeni'],
        [['']],
        col_widths=[270*mm]
    ))
    story.append(sp2)

    # Imza
    story.append(section_banner('IMZA BLOGU'))
    story.append(sp)
    story.append(styled_table(
        ['', 'Hazirlayan (Operator)', 'Kontrol Eden (Kalite)', 'Onaylayan (Uretim Amiri)'],
        [
            ['Ad Soyad', '', '', ''],
            ['Imza', '', '', ''],
            ['Tarih', '', '', ''],
        ],
        col_widths=[50*mm, 70*mm, 70*mm, 70*mm]
    ))

    # Footer note
    story.append(sp2)
    story.append(p('Bu dokuman ISO 9001:2015 Kalite Yonetim Sistemi kapsaminda hazirlanmistir. Kontrolsuz kopyasi gecersizdir.', FOOTER_S))

    doc.build(story)
    print(f'PDF olusturuldu: {OUTPUT}')


if __name__ == '__main__':
    build_report()

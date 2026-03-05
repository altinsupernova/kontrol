#!/usr/bin/env python3
"""Proses Kontrol Raporu - Google Sheets uyumlu Excel dosyası oluşturucu"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

OUTPUT = '/home/user/kontrol/rapor/PROSES_KONTROL_RAPORU.xlsx'

# Stiller
DARK_BLUE_FILL = PatternFill(start_color='29417A', end_color='29417A', fill_type='solid')
LIGHT_BLUE_FILL = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')
ALT_ROW_FILL = PatternFill(start_color='F0F5FA', end_color='F0F5FA', fill_type='solid')
WHITE_FONT = Font(name='Arial', size=9, bold=True, color='FFFFFF')
BOLD_FONT = Font(name='Arial', size=9, bold=True)
NORMAL_FONT = Font(name='Arial', size=9)
TITLE_FONT = Font(name='Arial', size=14, bold=True)
SECTION_FONT = Font(name='Arial', size=11, bold=True, color='FFFFFF')
SUBSECTION_FONT = Font(name='Arial', size=10, bold=True)
SMALL_FONT = Font(name='Arial', size=8, italic=True, color='666666')
THIN_BORDER = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)


def write_section_header(ws, row, text, cols=4):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    cell.fill = DARK_BLUE_FILL
    cell.alignment = LEFT
    return row + 1


def write_subsection(ws, row, text, cols=4):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SUBSECTION_FONT
    cell.fill = LIGHT_BLUE_FILL
    cell.alignment = LEFT
    return row + 1


def write_header_row(ws, row, headers):
    for ci, h in enumerate(headers, 1):
        cell = ws.cell(row=row, column=ci, value=h)
        cell.font = WHITE_FONT
        cell.fill = DARK_BLUE_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    return row + 1


def write_data_row(ws, row, values, bold_first=False, alt=False):
    for ci, v in enumerate(values, 1):
        cell = ws.cell(row=row, column=ci, value=v)
        cell.font = BOLD_FONT if (bold_first and ci == 1) else NORMAL_FONT
        cell.alignment = CENTER if ci > 1 else LEFT
        cell.border = THIN_BORDER
        if alt:
            cell.fill = ALT_ROW_FILL
    return row + 1


def write_info_row(ws, row, label, value, col_start=1):
    c1 = ws.cell(row=row, column=col_start, value=label)
    c1.font = BOLD_FONT
    c1.fill = ALT_ROW_FILL
    c1.border = THIN_BORDER
    c1.alignment = LEFT
    c2 = ws.cell(row=row, column=col_start + 1, value=value)
    c2.font = NORMAL_FONT
    c2.border = THIN_BORDER
    c2.alignment = LEFT


def build():
    wb = Workbook()

    # ═══════════ SAYFA 1: GENEL + PARAMETRELER ═══════════
    ws = wb.active
    ws.title = 'A - Genel Bilgiler'
    ws.sheet_properties.tabColor = '29417A'

    # Sütun genişlikleri
    for col, w in [(1, 22), (2, 25), (3, 22), (4, 25)]:
        ws.column_dimensions[chr(64 + col)].width = w

    # Başlık
    ws.merge_cells('A1:D1')
    c = ws.cell(row=1, column=1, value='PLASTİK ENJEKSİYON PROSES KONTROL RAPORU')
    c.font = TITLE_FONT
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:D2')
    c = ws.cell(row=2, column=1, value='Doküman No: KKR-ENJ-001  |  Rev: 01  |  Ref: ISO 9001:2015 Madde 8.5.1 / ISO 2859-1 (AQL)')
    c.font = Font(name='Arial', size=8, color='555555')
    c.alignment = Alignment(horizontal='center')

    r = 4
    r = write_section_header(ws, r, 'BÖLÜM A: GENEL BİLGİLER')
    r += 1

    info = [
        ('Müşteri', 'EBERLE', 'Ürün Kodu', ''),
        ('Ürün Adı', 'KARE KCK K.M', 'Sipariş No', ''),
        ('Sipariş Adedi', '100.000 Adet', 'Parti / Lot No', ''),
        ('Makine No / Adı', 'BORCHE (Proses 6)', 'Hammadde Parti No', ''),
        ('Kalıp No', 'EBERLE.KARE.KCK.K.M', 'Rapor Tarihi', '…/…/202…'),
        ('Hammadde Tipi', '', 'Kontrol Sıklığı', 'Her 50 Baskı'),
    ]
    for label1, val1, label2, val2 in info:
        write_info_row(ws, r, label1, val1, 1)
        write_info_row(ws, r, label2, val2, 3)
        r += 1

    # ═══════════ SAYFA 2: MAKİNE PARAMETRELERİ ═══════════
    ws2 = wb.create_sheet('B - Makine Param.')
    ws2.sheet_properties.tabColor = '29417A'
    for col, w in [(1, 12), (2, 20), (3, 18), (4, 18), (5, 18)]:
        ws2.column_dimensions[chr(64 + col)].width = w

    r = 1
    r = write_section_header(ws2, r, 'BÖLÜM B: MAKİNE PARAMETRELERİ', 5)
    r += 1

    # B.1
    r = write_subsection(ws2, r, 'B.1 — Enjeksiyon Parametreleri (4 Aşama)', 5)
    r = write_header_row(ws2, r, ['Aşama', 'Basınç (bar)', 'Hız (%)', 'Pozisyon (mm)', ''])
    for vals in [['1','161.0','80.0','11.0',''], ['2','155.0','45.0','10.0',''], ['3','140.0','40.0','9.0',''], ['4','160.0','35.0','—','']]:
        alt = vals[0] in ('2', '4')
        r = write_data_row(ws2, r, vals, alt=alt)
    r += 1

    # B.2
    r = write_subsection(ws2, r, 'B.2 — Tutma (Holding) Parametreleri (3 Aşama)', 5)
    r = write_header_row(ws2, r, ['Aşama', 'Basınç (bar)', 'Hız (%)', 'Süre (s)', ''])
    for vals in [['1','60.0','5.0','1.00',''], ['2','0.0','0.0','0.00',''], ['3','0.0','0.0','0.00','']]:
        alt = vals[0] == '2'
        r = write_data_row(ws2, r, vals, alt=alt)
    r += 1

    # B.3
    r = write_subsection(ws2, r, 'B.3 — Plastikleştirme (Mal Alma) (3 Aşama)', 5)
    r = write_header_row(ws2, r, ['Aşama', 'Geri Basınç (bar)', 'Basınç (bar)', 'Hız (%)', 'Pozisyon (mm)'])
    for vals in [['1','5.0','90','70','10.0'], ['2','8.0','90','70','15.0'], ['3','10.0','90','70','28.0']]:
        alt = vals[0] == '2'
        r = write_data_row(ws2, r, vals, alt=alt)
    r += 1

    # B.4
    r = write_subsection(ws2, r, 'B.4 — Sıcaklık Profili (Isıtma Bölgeleri)', 5)
    r = write_header_row(ws2, r, ['Bölge 1 (Nozül)', 'Bölge 2', 'Bölge 3', 'Bölge 4', 'Bölge 5'])
    r = write_data_row(ws2, r, ['275 °C', '270 °C', '270 °C', '271 °C', '270 °C'])
    r += 1
    r = write_header_row(ws2, r, ['Parametre', 'Değer', '', '', ''])
    r = write_data_row(ws2, r, ['Bekleme Isısı', '100 °C', '', '', ''], bold_first=True)
    r = write_data_row(ws2, r, ['Yağ Sıcaklığı', '21 °C', '', '', ''], bold_first=True, alt=True)
    r += 1

    # B.6
    r = write_subsection(ws2, r, 'B.6 — Ek Kesme / Zamanlama Parametreleri', 5)
    r = write_header_row(ws2, r, ['Parametre', 'Değer', '', '', ''])
    ek_params = [
        ['Enjek Basınç (Gerçek)', '159.4 – 159.5 bar'],
        ['Enjek Zaman', '4.00 s'],
        ['Yastıklama (Cushion)', '6.36 – 6.45 mm'],
        ['Anlık Tutma Zamanı', '1.0 s'],
        ['Cut-off Pozisyonu', '21.6 mm'],
        ['Vida Pozisyonu', '31.26 – 31.30 mm'],
    ]
    for i, (p, v) in enumerate(ek_params):
        r = write_data_row(ws2, r, [p, v, '', '', ''], bold_first=True, alt=(i % 2 == 1))

    # ═══════════ SAYFA 3: B.5 ÖZET (geniş tablo - ayrı sayfa) ═══════════
    ws3 = wb.create_sheet('B.5 - Özet Param.')
    ws3.sheet_properties.tabColor = '4472C4'
    for col, w in enumerate([5, 22, 8, 12, 12, 12, 14, 14, 14, 10], 1):
        ws3.column_dimensions[chr(64 + col) if col <= 9 else ('A' + chr(64 + col - 9) if col == 10 else chr(64+col))].width = w
    # Fix column J width
    ws3.column_dimensions['A'].width = 5
    ws3.column_dimensions['B'].width = 22
    ws3.column_dimensions['C'].width = 8
    ws3.column_dimensions['D'].width = 12
    ws3.column_dimensions['E'].width = 12
    ws3.column_dimensions['F'].width = 12
    ws3.column_dimensions['G'].width = 14
    ws3.column_dimensions['H'].width = 14
    ws3.column_dimensions['I'].width = 14
    ws3.column_dimensions['J'].width = 10

    r = 1
    r = write_section_header(ws3, r, 'B.5 — Özet Proses Parametreleri', 10)
    r += 1
    r = write_header_row(ws3, r, ['No', 'Parametre', 'Birim', 'Alt Limit', 'Hedef Değer', 'Üst Limit', 'Gerçekleşen 1', 'Gerçekleşen 2', 'Gerçekleşen 3', 'Durum'])

    ozet = [
        ['1', 'Kalıp Sıcaklığı', '°C', '—', '—', '—', '—', '—', '—', '—'],
        ['2', 'Eriyik Sıcaklığı', '°C', '270', '272', '275', '', '', '', ''],
        ['3', 'Enj. Basınç (Aş.1)', 'bar', '155', '161', '165', '161.0', '161.0', '', ''],
        ['4', 'Enj. Hız (Aş.1)', '%', '75', '80', '85', '80.0', '80.0', '', ''],
        ['5', 'Ütüleme Basıncı', 'bar', '55', '60', '65', '60.0', '60.0', '', ''],
        ['6', 'Ütüleme Süresi', 's', '0.8', '1.0', '1.2', '1.00', '1.00', '', ''],
        ['7', 'Soğutma Süresi', 's', '23', '25', '27', '25.0', '25.0', '', ''],
        ['8', 'Çevrim Süresi', 's', '25', '27.6', '35', '27.6', '49.7*', '', ''],
        ['9', 'Sıkma Kuvveti', 'ton', '—', '—', '—', '—', '—', '—', '—'],
    ]
    for vals in ozet:
        alt = int(vals[0]) % 2 == 0 if vals[0].isdigit() else False
        r = write_data_row(ws3, r, vals, alt=alt)

    r += 1
    c = ws3.cell(row=r, column=1, value='* Çevrim süresi 49.7s: Olağandışı yüksek — araştırılmalı (duruş, kalıp açma gecikmesi vb.)')
    c.font = SMALL_FONT

    # ═══════════ SAYFA 4: KALİTE KONTROL ═══════════
    ws4 = wb.create_sheet('C - Kalite Kontrol')
    ws4.sheet_properties.tabColor = '548235'
    for col_letter, w in [('A',5),('B',14),('C',18),('D',14),('E',16),('F',14),('G',16),('H',16),('I',14),('J',18),('K',14),('L',16)]:
        ws4.column_dimensions[col_letter].width = w

    r = 1
    r = write_section_header(ws4, r, 'BÖLÜM C: GÖRSEL VE BOYUTSAL KALİTE KONTROL (ISO 2859-1 AQL)', 12)
    r += 1
    r = write_header_row(ws4, r, ['No', 'Kontrol Saati', 'Üretim Miktarı', 'Numune Adedi', 'Çöküntü İzi', 'Çapak', 'Çizik/Leke', 'Kalıp Yüzeyi', 'Renk Uyumu', 'Boyutsal Kontrol', 'Genel Sonuç', 'Operatör İmza'])

    for i in range(1, 13):
        vals = [str(i)] + [''] * 11
        r = write_data_row(ws4, r, vals, alt=(i % 2 == 0))

    # ═══════════ SAYFA 5: ÜRETİM ÖZETİ + HATA + KARAR ═══════════
    ws5 = wb.create_sheet('D-E-F Özet & Onay')
    ws5.sheet_properties.tabColor = 'C00000'
    ws5.column_dimensions['A'].width = 30
    ws5.column_dimensions['B'].width = 25
    ws5.column_dimensions['C'].width = 14
    ws5.column_dimensions['D'].width = 14
    ws5.column_dimensions['E'].width = 18
    ws5.column_dimensions['F'].width = 25

    r = 1
    r = write_section_header(ws5, r, 'BÖLÜM D: ÜRETİM ÖZETİ', 6)
    r += 1
    r = write_header_row(ws5, r, ['Parametre', 'Değer', '', '', '', ''])
    for i, (p, v) in enumerate([
        ('Toplam Üretim (Adet)', ''),
        ('Sevk Edilen (Adet)', ''),
        ('Hatalı / Fire (Adet)', ''),
        ('Fire Oranı (%)', '0,00%'),
        ('Kabul Oranı (%)', '0,00%'),
        ('AQL Seviyesi', '1.0 (Genel Kontrol Seviyesi II)'),
    ]):
        r = write_data_row(ws5, r, [p, v, '', '', '', ''], bold_first=True, alt=(i % 2 == 1))

    r += 2
    r = write_section_header(ws5, r, 'BÖLÜM E: HATA TÜRLERİ ANALİZİ', 6)
    r += 1
    r = write_header_row(ws5, r, ['No', 'Hata Türü', 'Hata Adedi', 'Oran (%)', 'Sınıflandırma', 'Düzeltici Faaliyet'])
    hatalar = [
        ('1', 'Çöküntü İzi (Sink Mark)', '', '0,0%', '', ''),
        ('2', 'Çapak (Flash)', '', '0,0%', '', ''),
        ('3', 'Kısa Baskı (Short Shot)', '', '0,0%', '', ''),
        ('4', 'Yanık İzi (Burn Mark)', '', '0,0%', '', ''),
        ('5', 'Eğilme / Çarpılma (Warpage)', '', '0,0%', '', ''),
        ('6', 'Çizik / Leke', '', '0,0%', '', ''),
        ('7', 'Renk Uyumsuzluğu', '', '0,0%', '', ''),
        ('8', 'Diğer', '', '0,0%', '', ''),
        ('', 'TOPLAM HATA', '0', '0,0%', '', ''),
    ]
    for i, vals in enumerate(hatalar):
        r = write_data_row(ws5, r, list(vals), bold_first=(i == len(hatalar)-1), alt=(i % 2 == 1))

    r += 2
    r = write_section_header(ws5, r, 'BÖLÜM F: KARAR VE ONAY', 6)
    r += 1
    c = ws5.cell(row=r, column=1, value='PARTİ KARARI:')
    c.font = BOLD_FONT
    ws5.cell(row=r, column=2, value='☐ KABUL').font = BOLD_FONT
    ws5.cell(row=r, column=3, value='☐ ŞARTLI KABUL').font = BOLD_FONT
    ws5.cell(row=r, column=4, value='☐ RED').font = BOLD_FONT
    for ci in range(1, 5):
        ws5.cell(row=r, column=ci).border = THIN_BORDER
    r += 2

    c = ws5.cell(row=r, column=1, value='Açıklama / Şartlı Kabul Nedeni:')
    c.font = BOLD_FONT
    ws5.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws5.cell(row=r, column=2).border = THIN_BORDER
    r += 3

    # İmza
    r = write_section_header(ws5, r, 'İMZA BLOĞU', 6)
    r += 1
    r = write_header_row(ws5, r, ['', 'Hazırlayan (Operatör)', 'Kontrol Eden (Kalite)', 'Onaylayan (Üretim Amiri)', '', ''])
    for label in ['Ad Soyad', 'İmza', 'Tarih']:
        r = write_data_row(ws5, r, [label, '', '', '', '', ''], bold_first=True)

    r += 2
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    c = ws5.cell(row=r, column=1, value='Bu doküman ISO 9001:2015 Kalite Yönetim Sistemi kapsamında hazırlanmıştır. Kontrolsüz kopyası geçersizdir.')
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal='center')

    wb.save(OUTPUT)
    print(f'Excel dosyası oluşturuldu: {OUTPUT}')


if __name__ == '__main__':
    build()

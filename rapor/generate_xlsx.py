#!/usr/bin/env python3
"""Proses Kontrol Raporu - TEK SAYFA - Google Sheets uyumlu Excel"""

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
RED_FONT = Font(name='Arial', size=8, italic=True, color='CC0000')
THIN_BORDER = Border(
    left=Side(style='thin', color='CCCCCC'),
    right=Side(style='thin', color='CCCCCC'),
    top=Side(style='thin', color='CCCCCC'),
    bottom=Side(style='thin', color='CCCCCC')
)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center', wrap_text=True)

COLS = 12  # Toplam sütun sayısı


def section_header(ws, row, text):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=COLS)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SECTION_FONT
    cell.fill = DARK_BLUE_FILL
    cell.alignment = LEFT
    for c in range(1, COLS + 1):
        ws.cell(row=row, column=c).fill = DARK_BLUE_FILL
    return row + 1


def subsection(ws, row, text, span=None):
    span = span or COLS
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = SUBSECTION_FONT
    cell.fill = LIGHT_BLUE_FILL
    cell.alignment = LEFT
    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = LIGHT_BLUE_FILL
    return row + 1


def header_row(ws, row, headers, col_start=1):
    for ci, h in enumerate(headers):
        cell = ws.cell(row=row, column=col_start + ci, value=h)
        cell.font = WHITE_FONT
        cell.fill = DARK_BLUE_FILL
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    return row + 1


def data_row(ws, row, values, col_start=1, bold_first=False, alt=False):
    for ci, v in enumerate(values):
        cell = ws.cell(row=row, column=col_start + ci, value=v)
        cell.font = BOLD_FONT if (bold_first and ci == 0) else NORMAL_FONT
        cell.alignment = CENTER if ci > 0 else LEFT
        cell.border = THIN_BORDER
        if alt:
            cell.fill = ALT_ROW_FILL
    return row + 1


def info_pair(ws, row, label, value, col_start=1):
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
    ws = wb.active
    ws.title = 'Proses Kontrol Raporu'
    ws.sheet_properties.tabColor = '29417A'

    # Sütun genişlikleri (A-L = 12 sütun)
    widths = {'A': 5, 'B': 20, 'C': 16, 'D': 14, 'E': 14, 'F': 14,
              'G': 14, 'H': 14, 'I': 14, 'J': 16, 'K': 14, 'L': 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    # ═══════════ BAŞLIK ═══════════
    ws.merge_cells(f'A1:L1')
    c = ws.cell(row=1, column=1, value='PLASTİK ENJEKSİYON PROSES KONTROL RAPORU')
    c.font = TITLE_FONT
    c.alignment = Alignment(horizontal='center')

    ws.merge_cells(f'A2:L2')
    c = ws.cell(row=2, column=1, value='Doküman No: KKR-ENJ-001  |  Rev: 01  |  Ref: ISO 9001:2015 Madde 8.5.1 / ISO 2859-1 (AQL)')
    c.font = Font(name='Arial', size=8, color='555555')
    c.alignment = Alignment(horizontal='center')

    # ═══════════ BÖLÜM A: GENEL BİLGİLER ═══════════
    r = 4
    r = section_header(ws, r, 'BÖLÜM A: GENEL BİLGİLER')

    info = [
        ('Müşteri', 'EBERLE', 'Ürün Kodu', '', 'Sipariş Adedi', '100.000 Adet'),
        ('Ürün Adı', 'KARE KCK K.M', 'Sipariş No', '', 'Parti / Lot No', ''),
        ('Makine No / Adı', 'BORCHE (Proses 6)', 'Kalıp No', 'EBERLE.KARE.KCK.K.M', 'Hammadde Parti No', ''),
        ('Hammadde Tipi', '', 'Rapor Tarihi', '…/…/202…', 'Kontrol Sıklığı', 'Her 50 Baskı'),
    ]
    for row_data in info:
        info_pair(ws, r, row_data[0], row_data[1], 1)
        info_pair(ws, r, row_data[2], row_data[3], 3)
        info_pair(ws, r, row_data[4], row_data[5], 5)
        r += 1

    # ═══════════ BÖLÜM B: MAKİNE PARAMETRELERİ ═══════════
    r += 1
    r = section_header(ws, r, 'BÖLÜM B: MAKİNE PARAMETRELERİ')

    # B.1 Enjeksiyon (sol) + B.2 Tutma (sağ) yan yana
    r = subsection(ws, r, 'B.1 — Enjeksiyon Parametreleri                                                              B.2 — Tutma (Holding) Parametreleri')
    r = header_row(ws, r, ['Aşama', 'Basınç (bar)', 'Hız (%)', 'Pozisyon (mm)'], col_start=1)
    # B.2 header aynı satıra
    for ci, h in enumerate(['', 'Aşama', 'Basınç (bar)', 'Hız (%)', 'Süre (s)']):
        if h:
            cell = ws.cell(row=r - 1, column=5 + ci, value=h)
            cell.font = WHITE_FONT
            cell.fill = DARK_BLUE_FILL
            cell.alignment = CENTER
            cell.border = THIN_BORDER

    enj = [['1','161.0','80.0','11.0'], ['2','155.0','45.0','10.0'], ['3','140.0','40.0','9.0'], ['4','160.0','35.0','—']]
    tut = [['1','60.0','5.0','1.00'], ['2','0.0','0.0','0.00'], ['3','0.0','0.0','0.00']]

    for i in range(4):
        for ci, v in enumerate(enj[i]):
            cell = ws.cell(row=r, column=1 + ci, value=v)
            cell.font = NORMAL_FONT
            cell.alignment = CENTER
            cell.border = THIN_BORDER
            if i % 2 == 1:
                cell.fill = ALT_ROW_FILL
        # B.2 - sağ taraf
        if i < 3:
            for ci, v in enumerate(tut[i]):
                cell = ws.cell(row=r, column=6 + ci, value=v)
                cell.font = NORMAL_FONT
                cell.alignment = CENTER
                cell.border = THIN_BORDER
                if i % 2 == 1:
                    cell.fill = ALT_ROW_FILL
        r += 1

    # B.3 Plastikleştirme
    r += 1
    r = subsection(ws, r, 'B.3 — Plastikleştirme (Mal Alma) (3 Aşama)')
    r = header_row(ws, r, ['Aşama', 'Geri Basınç (bar)', 'Basınç (bar)', 'Hız (%)', 'Pozisyon (mm)'], col_start=1)
    for vals in [['1','5.0','90','70','10.0'], ['2','8.0','90','70','15.0'], ['3','10.0','90','70','28.0']]:
        alt = vals[0] == '2'
        r = data_row(ws, r, vals, alt=alt)

    # B.4 Sıcaklık + B.6 Ek Parametreler yan yana
    r += 1
    r = subsection(ws, r, 'B.4 — Sıcaklık Profili                                                                                B.6 — Ek Kesme / Zamanlama')
    r = header_row(ws, r, ['Bölge 1 (Nozül)', 'Bölge 2', 'Bölge 3', 'Bölge 4', 'Bölge 5'], col_start=1)
    # B.6 başlık sağ
    for ci, h in enumerate(['', 'Parametre', 'Değer']):
        if h:
            cell = ws.cell(row=r - 1, column=6 + ci, value=h)
            cell.font = WHITE_FONT
            cell.fill = DARK_BLUE_FILL
            cell.alignment = CENTER
            cell.border = THIN_BORDER

    # Sıcaklık verileri + sağda B.6
    sicaklik = ['275 °C', '270 °C', '270 °C', '271 °C', '270 °C']
    ek_params = [
        ('Enjek Basınç (Gerçek)', '159.4 – 159.5 bar'),
        ('Enjek Zaman', '4.00 s'),
        ('Yastıklama (Cushion)', '6.36 – 6.45 mm'),
        ('Anlık Tutma Zamanı', '1.0 s'),
        ('Cut-off Pozisyonu', '21.6 mm'),
        ('Vida Pozisyonu', '31.26 – 31.30 mm'),
    ]

    # Satır 1: sıcaklık + ilk ek param
    for ci, v in enumerate(sicaklik):
        cell = ws.cell(row=r, column=1 + ci, value=v)
        cell.font = NORMAL_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    for ei, (p, v) in enumerate(ek_params):
        row_idx = r + ei
        c1 = ws.cell(row=row_idx, column=7, value=p)
        c1.font = BOLD_FONT
        c1.alignment = LEFT
        c1.border = THIN_BORDER
        if ei % 2 == 1:
            c1.fill = ALT_ROW_FILL
        c2 = ws.cell(row=row_idx, column=8, value=v)
        c2.font = NORMAL_FONT
        c2.alignment = CENTER
        c2.border = THIN_BORDER
        if ei % 2 == 1:
            c2.fill = ALT_ROW_FILL
    r += 1

    # Bekleme/Yağ sıcaklığı sol tarafa
    for ci, (p, v) in enumerate([('Bekleme Isısı', '100 °C'), ('Yağ Sıcaklığı', '21 °C')]):
        c1 = ws.cell(row=r, column=1, value=p)
        c1.font = BOLD_FONT
        c1.border = THIN_BORDER
        c1.alignment = LEFT
        if ci == 1:
            c1.fill = ALT_ROW_FILL
        c2 = ws.cell(row=r, column=2, value=v)
        c2.font = NORMAL_FONT
        c2.border = THIN_BORDER
        c2.alignment = CENTER
        if ci == 1:
            c2.fill = ALT_ROW_FILL
        r += 1

    # Kalan B.6 satırları zaten yukarıda dolduruldu, r'yi ek_params sonuna taşı
    r = max(r, r)  # zaten doğru
    r += 1

    # B.5 Özet Proses Parametreleri
    r = subsection(ws, r, 'B.5 — Özet Proses Parametreleri')
    r = header_row(ws, r, ['No', 'Parametre', 'Birim', 'Alt Limit', 'Hedef Değer', 'Üst Limit', 'Gerçekleşen 1', 'Gerçekleşen 2', 'Gerçekleşen 3', 'Durum'], col_start=1)
    ozet = [
        ['1', 'Kalıp Sıcaklığı', '°C', '—', '—', '—', '', '', '', ''],
        ['2', 'Eriyik Sıcaklığı', '°C', '270', '272', '275', '', '', '', ''],
        ['3', 'Enj. Basınç (Aş.1)', 'bar', '155', '161', '165', '161.0', '161.0', '', ''],
        ['4', 'Enj. Hız (Aş.1)', '%', '75', '80', '85', '80.0', '80.0', '', ''],
        ['5', 'Ütüleme Basıncı', 'bar', '55', '60', '65', '60.0', '60.0', '', ''],
        ['6', 'Ütüleme Süresi', 's', '0.8', '1.0', '1.2', '1.00', '1.00', '', ''],
        ['7', 'Soğutma Süresi', 's', '23', '25', '27', '25.0', '25.0', '', ''],
        ['8', 'Çevrim Süresi', 's', '25', '27.6', '35', '27.6', '49.7*', '', ''],
        ['9', 'Sıkma Kuvveti', 'ton', '—', '—', '—', '', '', '', ''],
    ]
    for vals in ozet:
        alt = int(vals[0]) % 2 == 0
        r = data_row(ws, r, vals, alt=alt)

    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=10)
    c = ws.cell(row=r, column=1, value='* Çevrim süresi 49.7s: Olağandışı yüksek — araştırılmalı')
    c.font = RED_FONT
    r += 1

    # ═══════════ BÖLÜM C: KALİTE KONTROL ═══════════
    r += 1
    r = section_header(ws, r, 'BÖLÜM C: GÖRSEL VE BOYUTSAL KALİTE KONTROL (ISO 2859-1 AQL)')
    r = header_row(ws, r, ['No', 'Kontrol Saati', 'Üretim Miktarı', 'Numune Adedi', 'Çöküntü İzi', 'Çapak', 'Çizik/Leke', 'Kalıp Yüzeyi', 'Renk Uyumu', 'Boyutsal Kontrol', 'Genel Sonuç', 'Operatör İmza'], col_start=1)
    for i in range(1, 13):
        vals = [str(i)] + [''] * 11
        r = data_row(ws, r, vals, alt=(i % 2 == 0))

    # ═══════════ BÖLÜM D: ÜRETİM ÖZETİ ═══════════
    r += 1
    r = section_header(ws, r, 'BÖLÜM D: ÜRETİM ÖZETİ')
    r = header_row(ws, r, ['Parametre', 'Değer'], col_start=1)
    for i, (p, v) in enumerate([
        ('Toplam Üretim (Adet)', ''),
        ('Sevk Edilen (Adet)', ''),
        ('Hatalı / Fire (Adet)', ''),
        ('Fire Oranı (%)', '0,00%'),
        ('Kabul Oranı (%)', '0,00%'),
        ('AQL Seviyesi', '1.0 (Genel Kontrol Seviyesi II)'),
    ]):
        r = data_row(ws, r, [p, v], bold_first=True, alt=(i % 2 == 1))

    # ═══════════ BÖLÜM E: HATA ANALİZİ ═══════════
    r += 1
    r = section_header(ws, r, 'BÖLÜM E: HATA TÜRLERİ ANALİZİ')
    r = header_row(ws, r, ['No', 'Hata Türü', 'Hata Adedi', 'Oran (%)', 'Sınıflandırma', 'Düzeltici Faaliyet'], col_start=1)
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
        r = data_row(ws, r, list(vals), bold_first=(i == len(hatalar) - 1), alt=(i % 2 == 1))

    # ═══════════ BÖLÜM F: KARAR VE ONAY ═══════════
    r += 1
    r = section_header(ws, r, 'BÖLÜM F: KARAR VE ONAY')

    ws.cell(row=r, column=1, value='PARTİ KARARI:').font = BOLD_FONT
    ws.cell(row=r, column=2, value='☐ KABUL').font = BOLD_FONT
    ws.cell(row=r, column=3, value='☐ ŞARTLI KABUL').font = BOLD_FONT
    ws.cell(row=r, column=4, value='☐ RED').font = BOLD_FONT
    for ci in range(1, 5):
        ws.cell(row=r, column=ci).border = THIN_BORDER
    r += 1

    ws.cell(row=r, column=1, value='Açıklama / Şartlı Kabul Nedeni:').font = BOLD_FONT
    ws.cell(row=r, column=1).border = THIN_BORDER
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    ws.cell(row=r, column=2).border = THIN_BORDER
    r += 2

    # İmza bloğu
    r = subsection(ws, r, 'İMZA BLOĞU')
    r = header_row(ws, r, ['', 'Hazırlayan (Operatör)', 'Kontrol Eden (Kalite)', 'Onaylayan (Üretim Amiri)'], col_start=1)
    for label in ['Ad Soyad', 'İmza', 'Tarih']:
        r = data_row(ws, r, [label, '', '', ''], bold_first=True)

    r += 1
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COLS)
    c = ws.cell(row=r, column=1, value='Bu doküman ISO 9001:2015 Kalite Yönetim Sistemi kapsamında hazırlanmıştır. Kontrolsüz kopyası geçersizdir.')
    c.font = SMALL_FONT
    c.alignment = Alignment(horizontal='center')

    # Yazdırma ayarları - tek sayfaya sığdır
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr = ws.sheet_properties.pageSetUpPr or None

    wb.save(OUTPUT)
    print(f'Excel dosyası oluşturuldu: {OUTPUT}')
    print(f'Toplam satır: {r}')


if __name__ == '__main__':
    build()

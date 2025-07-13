# Excel Data Comparator

Bu Python scripti, belirlenen Excel dosyalarını okuyup karşılaştırma yapar.  
Eşleşen ve eşleşmeyen verileri ayrı sayfalara yazar, farkları hesaplar ve renklerle işaretler.

---

## Script Açıklaması

- `config.json` dosyasından dosya yolları ve ayarlar okunur.
- Belirlenen Excel sayfalarında başlık satırları hedef satıra kaydırılır.
- Karşılaştırma dosyasından veriler pandas ile okunur ve gruplanır.
- Giriş dosyasındaki ve karşılaştırma dosyasındaki değerleri karşılaştırılır.
- Eşleşen ve eşleşmeyen değerler ayrı sayfalara yazılır.
- Eşleşenler için toplamlar hesaplanır, farklar bulunur ve belirli eşiklere göre renklendirilir.
- Eşleşmeyenler için hangi sayfada bulunduğu işaretlenir.

---

## Python Kodu

```python
import os
import re
import json
import openpyxl
import pandas as pd
from collections import defaultdict
from openpyxl.styles import PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

# Ortam değişkenlerini yükle
load_dotenv()

# Konfigürasyon dosyasını yükle
with open('config.json', 'r', encoding='utf-8') as f:
    config = json.load(f)

# Dosya yolları
input_file_path = config['input_file']
comparison_file_path = config['comparison_file']
output_file_path = config['output_file']

sheet_ids = [str(s) for s in config['sheet_ids']]
column_headers = config['column_headers']
target_row = config['headers_row_target']

# Excel dosyasını yükle
input_wb = openpyxl.load_workbook(input_file_path)

def shift_rows_down(ws, start_row, shift_count):
    """Satırları aşağı kaydırır."""
    ws.insert_rows(start_row, shift_count)

# Tüm sayfalar için başlıklar hedef satıra alınır
for sheet_name in sheet_ids:
    ws = input_wb[sheet_name]

    if ws[f"A{target_row}"].value != column_headers[0]:
        df = pd.DataFrame(ws.values)

        header_row_index = None
        for idx, row in df.iterrows():
            if row.tolist()[:len(column_headers)] == column_headers:
                header_row_index = idx + 1
                break

        if header_row_index is not None:
            shift_count = target_row - header_row_index
            if shift_count > 0:
                shift_rows_down(ws, header_row_index, shift_count)

            for r_idx, row in df.iloc[header_row_index - 1:].iterrows():
                for c_idx, val in enumerate(row):
                    ws.cell(row=target_row + r_idx, column=c_idx + 1, value=val)
        else:
            print(f"{sheet_name} sayfasında başlık bulunamadı.")
    else:
        print(f"{sheet_name} sayfasında başlık zaten doğru konumda.")

# Giriş dosyasını kaydet
input_wb.save(input_file_path)

# --- comparison.xlsx dosyasını oku ve Sheet1'e kopyala ---
comparison_wb = openpyxl.load_workbook(comparison_file_path)
comparison_sheet = comparison_wb.active  # genelde 'Sheet' ya da 'Sheet1'

# Sheet sayfasını pandas'a aktar
data = comparison_sheet.values
columns = next(data)  # ilk satır başlıklar
df = pd.DataFrame(data, columns=columns)

# Gerekli sütun adlarını kontrol et (örnek: 'SD NO', 'Sağlam Tabaka' vb.)
required_columns = ['SD NO', 'Sağlam Tabaka', 'Total Fire Tabaka', 'Cutstar Fire']
missing = [col for col in required_columns if col not in df.columns]
if missing:
    raise ValueError(f"Gerekli sütun(lar) eksik: {missing}")

# Grupla ve toplamları al
summary_df = df.groupby('SD NO').agg({
    'Sağlam Tabaka': 'sum',
    'Total Fire Tabaka': 'sum',
    'Cutstar Fire': 'sum'
}).reset_index()

# Gerekirse SD NO’yu sayıya çevir
summary_df['SD NO'] = summary_df['SD NO'].apply(lambda x: float(re.sub(r'[^\d.]', '', str(x))) if pd.notna(x) else x)

# Test çıktısı
print(summary_df.head())

from openpyxl import Workbook

# Giriş dosyasını tekrar yükle (önceki adımda kapatılmış olabilir)
input_wb = openpyxl.load_workbook(input_file_path)
input_values_set = set()

# Tüm sayfalarda A sütunundaki değerleri topla (örnek: 105–113)
for sheet_name in sheet_ids:
    ws = input_wb[sheet_name]
    for row in ws.iter_rows(min_row=27, min_col=1, max_col=1, values_only=True):
        val = row[0]
        if val is None or val == column_headers[0]:
            continue
        try:
            cleaned = float(re.sub(r"[^\d.]", "", str(val)))
            input_values_set.add(cleaned)
        except:
            continue

# comparison.xlsx'ten gelen SD NO değerleri
comparison_values = set(summary_df['SD NO'].dropna().astype(float).tolist())

# Eşleşen ve eşleşmeyenleri ayır
matched = sorted(list(input_values_set & comparison_values))
unmatched = sorted(list(input_values_set - comparison_values))

# Yeni bir output çalışma kitabı oluştur
output_wb = Workbook()
matched_ws = output_wb.active
matched_ws.title = "Matched"
unmatched_ws = output_wb.create_sheet(title="Unmatched")

# Eşleşen değerleri yaz
matched_ws.append(["Matched Values"])
for val in matched:
    matched_ws.append([val])

# Eşleşmeyen değerleri yaz
unmatched_ws.append(["Unmatched Values"])
for val in unmatched:
    unmatched_ws.append([val])

# Kaydet
output_wb.save(output_file_path)

print(f"{len(matched)} eşleşme ve {len(unmatched)} eşleşmeyen bulundu.")

# Output dosyasını tekrar aç
output_wb = openpyxl.load_workbook(output_file_path)
matched_ws = output_wb["Matched"]

# Başlıkları genişlet
matched_ws.cell(row=1, column=2, value="Comparison Total (Good+Waste)")
matched_ws.cell(row=1, column=3, value="Input Total (Good+Waste)")
matched_ws.cell(row=1, column=4, value="Difference")
matched_ws.cell(row=1, column=5, value="Source Sheet")

# Başlangıç verilerini al (2. satırdan itibaren)
matched_values = [matched_ws.cell(row=r, column=1).value for r in range(2, matched_ws.max_row + 1)]

# Karşılaştırma dictionary’si oluştur
comparison_dict = {
    float(row['SD NO']): row
    for _, row in summary_df.iterrows()
}

# Renkler
green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# Her matched değeri için input.xlsx sayfalarına bak
for i, val in enumerate(matched_values, start=2):
    comparison_brut = None
    input_brut = None
    found_in_sheet = None

    # Comparison dosyasından hesapla
    if val in comparison_dict:
        d = comparison_dict[val]
        try:
            comparison_brut = d['Sağlam Tabaka'] + d['Total Fire Tabaka']
        except:
            pass

    # Input dosyasından eşleşen satırı bul
    for sheet_name in sheet_ids:
        ws = input_wb[sheet_name]
        for row in ws.iter_rows(min_row=27, min_col=1, max_col=11):
            job_val = row[0].value
            if job_val is None:
                continue
            try:
                cleaned = float(re.sub(r"[^\d.]", "", str(job_val)))
            except:
                continue

            if cleaned == val:
                good = row[4].value  # E sütunu
                waste = row[5].value  # F sütunu
                if isinstance(good, (int, float)) and isinstance(waste, (int, float)):
                    input_brut = good + waste
                    found_in_sheet = sheet_name
                break
        if input_brut is not None:
            break  # bulunduysa sayfa aramasını bitir

    # Fark hesapla
    if isinstance(input_brut, (int, float)) and isinstance(comparison_brut, (int, float)):
        diff = input_brut - comparison_brut
    else:
        diff = None

    # Sonuçları yaz
    matched_ws.cell(row=i, column=2, value=comparison_brut)
    matched_ws.cell(row=i, column=3, value=input_brut)
    matched_ws.cell(row=i, column=4, value=diff)
    matched_ws.cell(row=i, column=5, value=found_in_sheet)

    # Renklendirme (Hücre veya satır)
    if isinstance(diff, (int, float)):
        if -200 <= diff <= 200:
            matched_ws.cell(row=i, column=4).fill = green_fill
        else:
            for col in range(1, 5):
                matched_ws.cell(row=i, column=col).fill = red_fill

# Sayıları bindelik formatla
for row in matched_ws.iter_rows(min_row=2, max_col=4):
    for cell in row:
        if isinstance(cell.value, (int, float)):
            cell.number_format = '#,##0'

# Kenarlıklar (isteğe bağlı)
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
for row in matched_ws.iter_rows():
    for cell in row:
        cell.border = thin_border

# Kaydet
output_wb.save(output_file_path)
output_wb.close()
input_wb.close()

print("Matched sayfasına detaylar işlendi ve farklar hesaplandı.")

# output.xlsx tekrar aç
output_wb = openpyxl.load_workbook(output_file_path)
unmatched_ws = output_wb["Unmatched"]

# Başlıkları genişlet
unmatched_ws.cell(row=1, column=2, value="Found In Sheet")

# Unmatched değerleri al
unmatched_values = [unmatched_ws.cell(row=r, column=1).value for r in range(2, unmatched_ws.max_row + 1)]

# input.xlsx tekrar yükle (önceki adımda kapatılmış olabilir)
input_wb = openpyxl.load_workbook(input_file_path)

# Her unmatched değer için input.xlsx sayfalarında ara
for i, val in enumerate(unmatched_values, start=2):
    found_sheet = "Not Found"
    for sheet_name in sheet_ids:
        ws = input_wb[sheet_name]
        for row in ws.iter_rows(min_row=27, min_col=1, max_col=1):
            job_val = row[0].value
            if job_val is None:
                continue
            try:
                cleaned = float(re.sub(r"[^\d.]", "", str(job_val)))
            except:
                continue

            if cleaned == val:
                found_sheet = sheet_name
                break
        if found_sheet != "Not Found":
            break  # bulunduysa diğer sayfalarda arama

    unmatched_ws.cell(row=i, column=2, value=found_sheet)

# İsteğe bağlı: stil ve renklendirme
gray_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
for row in range(2, unmatched_ws.max_row + 1):
    if unmatched_ws[f'B{row}'].value == "Not Found":
        for col in range(1, 3):
            unmatched_ws.cell(row=row, column=col).fill = gray_fill

# Kaydet ve kapat
output_wb.save(output_file_path)
output_wb.close()
input_wb.close()

print("Unmatched değerler etiketlendi.")

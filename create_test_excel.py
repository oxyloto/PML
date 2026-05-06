import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "CatalogData"

headers = [
    "GPART", "GPART DESC", "GTYPE",
    "GPART LEVEL", "COMMODITY", "COMMODITY DESC",
    "SCH", "SCH DESC",
    "SECT", "CATE", "CATE DESC", "SCOM"
]

header_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
header_font = Font(color="FFFFFF", bold=True)
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for col, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=col, value=h)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal='center')
    cell.border = border

# Senaryo 1-2: CATE=ALFBBB0, SCOM=ALFBBB0JJ mevcut, sadece GPART yeni
# Senaryo 3-4: CATE=ALFBBB0, SCOM=ALFBBB0KK yeni
# Senaryo 5: CATE=ALCCC0, SCOM=ALCCC0JJ tamamen yeni

rows = [
    # Senaryo: CATE EXIST + SCOM EXIST → sadece GPART oluşturulur
    # RecGpartLevel = GPART kodu - son 2 char (JJ)
    ["ALFBBB0BX00000000ZZJJ", "BLIND FLANGE ASME B16.47 SERIES B 150# RF ASTM A105 28.58MM x 28.58MM",
     "FLANGE", "ALFBBB0BX00000000ZZ",
     "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF",
     "ALFBBB0ZZ", "28.58MM",
     "PIPING-SECT", "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF", "ALFBBB0JJ"],

    ["ALFBBB0BX00000000ZZPP", "BLIND FLANGE ASME B16.47 SERIES B 150# RF ASTM A182 28.58MM x 28.58MM",
     "FLANGE", "ALFBBB0BX00000000ZZ",
     "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF",
     "ALFBBB0ZZ", "28.58MM",
     "PIPING-SECT", "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF", "ALFBBB0JJ"],

    # Senaryo: CATE EXIST + SCOM NEW → SCOM + GPART oluşturulur
    ["ALFBBB0BX00000000ZZKK", "BLIND FLANGE ASME B16.47 SERIES B 150# RF ASTM A105 33.32MM x 33.32MM",
     "FLANGE", "ALFBBB0BX00000000ZZ",
     "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF",
     "ALFBBB0KK", "33.32MM",
     "PIPING-SECT", "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF", "ALFBBB0KK"],

    ["ALFBBB0BX00000000KKPP", "BLIND FLANGE ASME B16.47 SERIES B 150# RF ASTM A182 33.32MM x 33.32MM",
     "FLANGE", "ALFBBB0BX00000000ZZ",
     "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF",
     "ALFBBB0KK", "33.32MM",
     "PIPING-SECT", "ALFBBB0", "BLIND FLANGE ASME B16.47 SERIES B 150# RF", "ALFBBB0KK"],

    # Senaryo: CATE NEW + SCOM NEW + GPART NEW → hepsi oluşturulur
    ["ALCCC0BX00000000ZZJJ", "BLIND FLANGE ASME B16.47 SERIES B 300# RF ASTM A105 28.58MM x 28.58MM",
     "FLANGE", "ALCCC0BX00000000ZZ",
     "ALCCC0", "BLIND FLANGE ASME B16.47 SERIES B 300# RF",
     "ALCCC0ZZ", "28.58MM",
     "PIPING-SECT", "ALCCC0", "BLIND FLANGE ASME B16.47 SERIES B 300# RF", "ALCCC0JJ"],
]

row_colors = ["FFF2CC", "FFF2CC", "D9E1F2", "D9E1F2", "E2EFDA"]

for i, (row_data, color) in enumerate(zip(rows, row_colors), 2):
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    for col, val in enumerate(row_data, 1):
        cell = ws.cell(row=i, column=col, value=val)
        cell.fill = fill
        cell.border = border
        cell.alignment = Alignment(horizontal='left')

for col in ws.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

ws.freeze_panes = "A2"

# Senaryo açıklamaları için ikinci sheet
ws2 = wb.create_sheet("Senaryo Aciklamalari")
ws2["A1"] = "Satır"
ws2["B1"] = "Beklenen Senaryo"
ws2["C1"] = "Yapılacak İşlem"
scenarios = [
    (2, "CATE=ALFBBB0 EXIST, SCOM=ALFBBB0JJ EXIST", "Sadece GPART oluşturulur"),
    (3, "CATE=ALFBBB0 EXIST, SCOM=ALFBBB0JJ EXIST", "Sadece GPART oluşturulur"),
    (4, "CATE=ALFBBB0 EXIST, SCOM=ALFBBB0KK NEW",   "SCOM + GPART oluşturulur"),
    (5, "CATE=ALFBBB0 EXIST, SCOM=ALFBBB0KK NEW",   "Sadece GPART oluşturulur (SCOM zaten 4. satırda yapıldı)"),
    (6, "CATE=ALCCC0 NEW, SCOM=ALCCC0JJ NEW",        "CATE + SCOM + GPART oluşturulur"),
]
for row_num, (satir, senaryo, islem) in enumerate(scenarios, 2):
    ws2.cell(row=row_num, column=1, value=satir)
    ws2.cell(row=row_num, column=2, value=senaryo)
    ws2.cell(row=row_num, column=3, value=islem)
for col in ws2.columns:
    max_len = max(len(str(c.value or "")) for c in col)
    ws2.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

wb.save("/home/user/PML/TEKCatalogGrid_TEST.xlsx")
print("OK")

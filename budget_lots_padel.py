"""
Génère un fichier Excel avec les scénarios de répartition
du budget lots pour le tournoi de Padel (84€).
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()

# --- Styles ---
title_font = Font(bold=True, size=14, color="FFFFFF")
header_font = Font(bold=True, size=11, color="FFFFFF")
bold_font = Font(bold=True, size=11)
currency_fmt = '#,##0.00 €'
pct_fmt = '0%'

fill_title = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
fill_header = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
fill_scenario_a = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
fill_scenario_b = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left = Alignment(horizontal='left', vertical='center', wrap_text=True)

BUDGET = 84

def style_range(ws, row, col_start, col_end, font=None, fill=None, alignment=None, number_format=None):
    for c in range(col_start, col_end + 1):
        cell = ws.cell(row=row, column=c)
        if font: cell.font = font
        if fill: cell.fill = fill
        if alignment: cell.alignment = alignment
        if number_format: cell.number_format = number_format
        cell.border = thin_border

def add_borders(ws, row_start, row_end, col_start, col_end):
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            ws.cell(row=r, column=c).border = thin_border

# ============================================================
# FEUILLE 1 : Scénarios de répartition
# ============================================================
ws1 = wb.active
ws1.title = "Scenarios"

# Titre
ws1.merge_cells('A1:G1')
ws1['A1'] = f"TOURNOI PADEL - Répartition budget lots ({BUDGET}€)"
ws1['A1'].font = title_font
ws1['A1'].fill = fill_title
ws1['A1'].alignment = center

ws1.merge_cells('A2:G2')
ws1['A2'] = "Différents scénarios selon le nombre de places récompensées et la valeur des lots"
ws1['A2'].font = Font(italic=True, size=10)
ws1['A2'].alignment = center

# --- Scénarios 2 places (1er + 2ème) ---
row = 4
ws1.merge_cells(f'A{row}:G{row}')
ws1[f'A{row}'] = "OPTION A : Récompenser 1er et 2ème uniquement"
ws1[f'A{row}'].font = Font(bold=True, size=12, color="2F5496")

row = 5
headers = ["Scénario", "1er (par pers.)", "1er (x2 joueurs)", "2ème (par pers.)", "2ème (x2 joueurs)", "Total", "Reste"]
for i, h in enumerate(headers, 1):
    ws1.cell(row=row, column=i, value=h)
style_range(ws1, row, 1, 7, font=header_font, fill=fill_header, alignment=center)

scenarios_2 = [
    ("A1 - Égalitaire",         21, 21),
    ("A2 - 60/40",              25.20, 16.80),
    ("A3 - 2/3 vs 1/3",        28, 14),
    ("A4 - Premium 1er",       30, 12),
    ("A5 - Gros écart",        35, 7),
]

for idx, (name, p1, p2) in enumerate(scenarios_2):
    r = row + 1 + idx
    ws1.cell(row=r, column=1, value=name).font = bold_font
    ws1.cell(row=r, column=2, value=p1).number_format = currency_fmt
    ws1.cell(row=r, column=3, value=p1*2).number_format = currency_fmt
    ws1.cell(row=r, column=4, value=p2).number_format = currency_fmt
    ws1.cell(row=r, column=5, value=p2*2).number_format = currency_fmt
    total = (p1 + p2) * 2
    ws1.cell(row=r, column=6, value=total).number_format = currency_fmt
    ws1.cell(row=r, column=7, value=BUDGET - total).number_format = currency_fmt
    fill = fill_scenario_a
    style_range(ws1, r, 1, 7, fill=fill, alignment=center)
    ws1.cell(row=r, column=1).alignment = left

add_borders(ws1, row, row + len(scenarios_2), 1, 7)

# --- Scénarios 3 places (1er + 2ème + 3ème) ---
row = row + len(scenarios_2) + 3
ws1.merge_cells(f'A{row}:G{row}')
ws1[f'A{row}'] = "OPTION B : Récompenser 1er, 2ème et 3ème"
ws1[f'A{row}'].font = Font(bold=True, size=12, color="548235")

row += 1
headers_3 = ["Scénario", "1er (par pers.)", "1er (x2)", "2ème (par pers.)", "2ème (x2)", "3ème (par pers.)", "3ème (x2)", "Total", "Reste"]
for i, h in enumerate(headers_3, 1):
    ws1.cell(row=row, column=i, value=h)
style_range(ws1, row, 1, 9, font=header_font, fill=PatternFill(start_color="548235", end_color="548235", fill_type="solid"), alignment=center)

scenarios_3 = [
    ("B1 - Égalitaire",       14, 14, 14),
    ("B2 - Dégressif doux",   18, 13, 11),
    ("B3 - Dégressif moyen",  20, 12, 10),
    ("B4 - Dégressif fort",   25, 10, 7),
    ("B5 - Premium 1er",      28, 8, 6),
    ("B6 - Symbolique 3ème",  22, 14, 6),
]

for idx, (name, p1, p2, p3) in enumerate(scenarios_3):
    r = row + 1 + idx
    ws1.cell(row=r, column=1, value=name).font = bold_font
    ws1.cell(row=r, column=2, value=p1).number_format = currency_fmt
    ws1.cell(row=r, column=3, value=p1*2).number_format = currency_fmt
    ws1.cell(row=r, column=4, value=p2).number_format = currency_fmt
    ws1.cell(row=r, column=5, value=p2*2).number_format = currency_fmt
    ws1.cell(row=r, column=6, value=p3).number_format = currency_fmt
    ws1.cell(row=r, column=7, value=p3*2).number_format = currency_fmt
    total = (p1 + p2 + p3) * 2
    ws1.cell(row=r, column=8, value=total).number_format = currency_fmt
    ws1.cell(row=r, column=9, value=BUDGET - total).number_format = currency_fmt
    style_range(ws1, r, 1, 9, fill=fill_scenario_b, alignment=center)
    ws1.cell(row=r, column=1).alignment = left

add_borders(ws1, row, row + len(scenarios_3), 1, 9)

# ============================================================
# FEUILLE 2 : Simulateur libre
# ============================================================
ws2 = wb.create_sheet("Simulateur")

ws2.merge_cells('A1:F1')
ws2['A1'] = "SIMULATEUR LIBRE - Entrez vos valeurs"
ws2['A1'].font = title_font
ws2['A1'].fill = fill_title
ws2['A1'].alignment = center

ws2['A3'] = "Budget total :"
ws2['A3'].font = bold_font
ws2['B3'] = BUDGET
ws2['B3'].number_format = currency_fmt
ws2['B3'].font = bold_font

ws2['A4'] = "Joueurs par équipe :"
ws2['A4'].font = bold_font
ws2['B4'] = 2
ws2['B4'].font = bold_font

row = 6
headers_sim = ["Place", "Valeur lot / personne", "Nb joueurs", "Sous-total", "% du budget"]
for i, h in enumerate(headers_sim, 1):
    ws2.cell(row=row, column=i, value=h)
style_range(ws2, row, 1, 5, font=header_font, fill=fill_header, alignment=center)

places = ["1er", "2ème", "3ème"]
for idx, place in enumerate(places):
    r = row + 1 + idx
    ws2.cell(row=r, column=1, value=place).font = bold_font
    ws2.cell(row=r, column=1).alignment = center
    # Valeur par personne (à remplir par l'utilisateur)
    ws2.cell(row=r, column=2, value=0).number_format = currency_fmt
    ws2.cell(row=r, column=2).alignment = center
    # Nb joueurs (formule liée à B4)
    ws2.cell(row=r, column=3, value=f'=$B$4')
    ws2.cell(row=r, column=3).alignment = center
    # Sous-total = valeur x nb joueurs
    ws2.cell(row=r, column=4).value = f'=B{r}*C{r}'
    ws2.cell(row=r, column=4).number_format = currency_fmt
    ws2.cell(row=r, column=4).alignment = center
    # % du budget
    ws2.cell(row=r, column=5).value = f'=IF($B$3=0,0,D{r}/$B$3)'
    ws2.cell(row=r, column=5).number_format = pct_fmt
    ws2.cell(row=r, column=5).alignment = center
    style_range(ws2, r, 1, 5, fill=PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"), alignment=center)

# Ligne Total
r_total = row + 4
ws2.cell(row=r_total, column=1, value="TOTAL").font = bold_font
ws2.cell(row=r_total, column=4).value = f'=SUM(D{row+1}:D{row+3})'
ws2.cell(row=r_total, column=4).number_format = currency_fmt
ws2.cell(row=r_total, column=5).value = f'=IF($B$3=0,0,D{r_total}/$B$3)'
ws2.cell(row=r_total, column=5).number_format = pct_fmt
style_range(ws2, r_total, 1, 5, fill=fill_total, alignment=center)

# Reste
r_reste = r_total + 1
ws2.cell(row=r_reste, column=1, value="RESTE").font = Font(bold=True, color="FF0000")
ws2.cell(row=r_reste, column=4).value = f'=$B$3-D{r_total}'
ws2.cell(row=r_reste, column=4).number_format = currency_fmt
ws2.cell(row=r_reste, column=4).font = Font(bold=True, color="FF0000")
style_range(ws2, r_reste, 1, 5, alignment=center)

add_borders(ws2, row, r_reste, 1, 5)

# Note
ws2.cell(row=r_reste + 2, column=1, value="Modifiez les cellules B7, B8, B9 pour simuler vos propres répartitions.").font = Font(italic=True, size=10, color="666666")

# ============================================================
# FEUILLE 3 : Idées de lots par gamme de prix
# ============================================================
ws3 = wb.create_sheet("Idees lots")

ws3.merge_cells('A1:C1')
ws3['A1'] = "IDEES DE LOTS PAR GAMME DE PRIX"
ws3['A1'].font = title_font
ws3['A1'].fill = fill_title
ws3['A1'].alignment = center

row = 3
headers_lots = ["Gamme de prix", "Idées de lots", "Remarque"]
for i, h in enumerate(headers_lots, 1):
    ws3.cell(row=row, column=i, value=h)
style_range(ws3, row, 1, 3, font=header_font, fill=fill_header, alignment=center)

lots = [
    ("5-10€",  "Surgrips, balles de padel (tube 3), bandeau/poignet, porte-clés padel", "Idéal pour 3ème place"),
    ("10-15€", "Chaussettes sport, bidon isotherme, serviette sport, balle anti-stress", "Bon compromis 2ème/3ème"),
    ("15-20€", "Sac à chaussures, casquette/visière, grip premium, protège-raquette", "Bon 2ème prix"),
    ("20-25€", "Pochette de raquette, t-shirt technique, lot balles + surgrips", "1er ou 2ème selon scénario"),
    ("25-35€", "Sac de padel compact, lunettes sport, abonnement magazine", "1er prix"),
]

for idx, (gamme, idees, remarque) in enumerate(lots):
    r = row + 1 + idx
    ws3.cell(row=r, column=1, value=gamme).font = bold_font
    ws3.cell(row=r, column=1).alignment = center
    ws3.cell(row=r, column=2, value=idees).alignment = left
    ws3.cell(row=r, column=3, value=remarque).alignment = left
    fill = fill_scenario_a if idx % 2 == 0 else PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    style_range(ws3, r, 1, 3, fill=fill)
    ws3.cell(row=r, column=1).alignment = center

add_borders(ws3, row, row + len(lots), 1, 3)

# --- Ajuster largeurs de colonnes ---
for ws in [ws1, ws2, ws3]:
    for col in range(1, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(col)].width = 22
    ws.column_dimensions['A'].width = 26

ws3.column_dimensions['B'].width = 60
ws3.column_dimensions['C'].width = 30

# Sauvegarder
output = "/home/user/planning-urbansoccer/budget_lots_padel.xlsx"
wb.save(output)
print(f"Fichier généré : {output}")

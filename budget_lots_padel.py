"""
Génère un fichier Excel pour le budget lots du tournoi de Padel (84€).
- Tableau gauche : catalogue des lots disponibles avec prix
- Tableau droit : attribution des lots par place (1er/2ème/3ème) avec listes déroulantes
"""
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Budget Lots"

# --- Styles ---
title_font = Font(bold=True, size=13, color="FFFFFF")
header_font = Font(bold=True, size=10, color="FFFFFF")
bold_font = Font(bold=True, size=10)
small_font = Font(size=10)
currency_fmt = '#,##0.00 €'

fill_title_blue = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
fill_title_green = PatternFill(start_color="548235", end_color="548235", fill_type="solid")
fill_header = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
fill_header_green = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
fill_1er = PatternFill(start_color="FFD966", end_color="FFD966", fill_type="solid")  # Or
fill_2eme = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")  # Argent
fill_3eme = PatternFill(start_color="DFC29A", end_color="DFC29A", fill_type="solid")  # Bronze
fill_total = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
fill_reste = PatternFill(start_color="FCE4EC", end_color="FCE4EC", fill_type="solid")
fill_light = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

center = Alignment(horizontal='center', vertical='center', wrap_text=True)
left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)

BUDGET = 84

def style_cell(cell, font=None, fill=None, alignment=None, number_format=None):
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if number_format: cell.number_format = number_format
    cell.border = thin_border

def style_row(ws, row, col_start, col_end, **kwargs):
    for c in range(col_start, col_end + 1):
        style_cell(ws.cell(row=row, column=c), **kwargs)

# ============================================================
# TABLEAU GAUCHE : Catalogue des lots (colonnes A-C)
# ============================================================
LEFT_START = 1  # col A
LEFT_END = 3    # col C

# Titre
ws.merge_cells('A1:C1')
c = ws['A1']
c.value = "CATALOGUE DES LOTS"
style_cell(c, font=title_font, fill=fill_title_blue, alignment=center)

# Budget
ws.merge_cells('A2:B2')
ws['A2'] = "Budget total :"
ws['A2'].font = bold_font
ws['A2'].alignment = Alignment(horizontal='right', vertical='center')
ws['C2'] = BUDGET
style_cell(ws['C2'], font=Font(bold=True, size=12, color="2F5496"), alignment=center, number_format=currency_fmt)

# En-têtes
row_header = 4
headers_left = ["#", "Lot", "Prix unitaire"]
for i, h in enumerate(headers_left):
    ws.cell(row=row_header, column=LEFT_START + i, value=h)
style_row(ws, row_header, LEFT_START, LEFT_END, font=header_font, fill=fill_header, alignment=center)

# Catalogue de lots (exemples à personnaliser)
lots = [
    ("Tube balles Padel (x3)", 8),
    ("Surgrips (lot de 3)", 6),
    ("Bandeau sport", 5),
    ("Bidon isotherme", 12),
    ("Chaussettes sport", 10),
    ("Serviette microfibre", 9),
    ("Casquette/Visière", 15),
    ("Sac à chaussures", 18),
    ("Protège-raquette", 16),
    ("T-shirt technique", 22),
    ("Pochette de raquette", 25),
    ("Sac de padel compact", 30),
]

lot_start_row = row_header + 1
for idx, (nom, prix) in enumerate(lots):
    r = lot_start_row + idx
    fill = fill_light if idx % 2 == 0 else fill_white
    ws.cell(row=r, column=1, value=idx + 1)
    style_cell(ws.cell(row=r, column=1), font=small_font, fill=fill, alignment=center)
    ws.cell(row=r, column=2, value=nom)
    style_cell(ws.cell(row=r, column=2), font=small_font, fill=fill, alignment=left_align)
    ws.cell(row=r, column=3, value=prix)
    style_cell(ws.cell(row=r, column=3), font=small_font, fill=fill, alignment=center, number_format=currency_fmt)

lot_end_row = lot_start_row + len(lots) - 1

# Plage nommée pour les lots (colonne B)
lot_names_range = f"'Budget Lots'!$B${lot_start_row}:$B${lot_end_row}"

# ============================================================
# TABLEAU DROIT : Attribution par place (colonnes E-G)
# ============================================================
RIGHT_COL = 5   # col E
RIGHT_END = 7   # col G

# Titre
ws.merge_cells('E1:G1')
c = ws['E1']
c.value = "ATTRIBUTION DES LOTS"
style_cell(c, font=title_font, fill=fill_title_green, alignment=center)

# Budget rappel
ws.merge_cells('E2:F2')
ws['E2'] = "Budget :"
ws['E2'].font = bold_font
ws['E2'].alignment = Alignment(horizontal='right', vertical='center')
ws['G2'] = f'=C2'
style_cell(ws['G2'], font=Font(bold=True, size=12, color="548235"), alignment=center, number_format=currency_fmt)

# En-têtes
headers_right = ["Place", "Lot attribué", "Coût"]
for i, h in enumerate(headers_right):
    ws.cell(row=row_header, column=RIGHT_COL + i, value=h)
style_row(ws, row_header, RIGHT_COL, RIGHT_END, font=header_font, fill=fill_header_green, alignment=center)

# --- Validation : liste déroulante basée sur la colonne B des lots ---
# On crée la formule de validation avec la plage des noms de lots
dv = DataValidation(
    type="list",
    formula1=f"=$B${lot_start_row}:$B${lot_end_row}",
    allow_blank=True
)
dv.error = "Choisissez un lot du catalogue"
dv.errorTitle = "Lot invalide"
dv.prompt = "Sélectionnez un lot du catalogue"
dv.promptTitle = "Choix du lot"
ws.add_data_validation(dv)

# Structure : 1er (4 lignes), 2ème (3 lignes), 3ème (2 lignes)
places = [
    ("1er", 4, fill_1er),
    ("2ème", 3, fill_2eme),
    ("3ème", 2, fill_3eme),
]

current_row = row_header + 1
attribution_rows = []  # pour collecter toutes les lignes de coût

for place_name, nb_lots, fill_place in places:
    for i in range(nb_lots):
        r = current_row
        attribution_rows.append(r)

        # Place (merge ou label uniquement sur la 1ère ligne)
        if i == 0:
            label = f"{place_name} (x{nb_lots} lots max)"
        else:
            label = ""
        ws.cell(row=r, column=RIGHT_COL, value=label)
        style_cell(ws.cell(row=r, column=RIGHT_COL), font=bold_font, fill=fill_place, alignment=center)

        # Lot attribué (dropdown)
        lot_cell = ws.cell(row=r, column=RIGHT_COL + 1)
        lot_cell.value = None
        style_cell(lot_cell, font=small_font, fill=fill_place, alignment=left_align)
        dv.add(lot_cell)

        # Coût = VLOOKUP sur le catalogue
        cost_cell = ws.cell(row=r, column=RIGHT_COL + 2)
        # RECHERCHEV : cherche le nom du lot dans B:C et retourne le prix (col 2)
        cost_cell.value = f'=IF(F{r}="",0,VLOOKUP(F{r},$B${lot_start_row}:$C${lot_end_row},2,FALSE))'
        style_cell(cost_cell, font=small_font, fill=fill_place, alignment=center, number_format=currency_fmt)

        current_row += 1

# Ligne séparatrice + sous-totaux par place
subtotal_row = current_row
# On calcule les plages pour chaque place
row_cursor = row_header + 1
for place_name, nb_lots, fill_place in places:
    start_r = row_cursor
    end_r = row_cursor + nb_lots - 1
    row_cursor = end_r + 1

# --- Ligne TOTAL ---
total_row = current_row
ws.cell(row=total_row, column=RIGHT_COL, value="TOTAL")
style_cell(ws.cell(row=total_row, column=RIGHT_COL), font=Font(bold=True, size=11), fill=fill_total, alignment=center)
ws.cell(row=total_row, column=RIGHT_COL + 1, value="")
style_cell(ws.cell(row=total_row, column=RIGHT_COL + 1), fill=fill_total, alignment=center)

first_cost = row_header + 1
last_cost = current_row - 1
ws.cell(row=total_row, column=RIGHT_END).value = f'=SUM(G{first_cost}:G{last_cost})'
style_cell(ws.cell(row=total_row, column=RIGHT_END), font=Font(bold=True, size=11), fill=fill_total, alignment=center, number_format=currency_fmt)

# --- Ligne RESTE ---
reste_row = total_row + 1
ws.cell(row=reste_row, column=RIGHT_COL, value="RESTE")
style_cell(ws.cell(row=reste_row, column=RIGHT_COL), font=Font(bold=True, size=11, color="FF0000"), fill=fill_reste, alignment=center)
ws.cell(row=reste_row, column=RIGHT_COL + 1, value="")
style_cell(ws.cell(row=reste_row, column=RIGHT_COL + 1), fill=fill_reste, alignment=center)
ws.cell(row=reste_row, column=RIGHT_END).value = f'=G2-G{total_row}'
style_cell(ws.cell(row=reste_row, column=RIGHT_END), font=Font(bold=True, size=11, color="FF0000"), fill=fill_reste, alignment=center, number_format=currency_fmt)

# --- Sous-totaux par place (en dessous) ---
detail_row = reste_row + 2
ws.cell(row=detail_row, column=RIGHT_COL, value="Détail par place")
style_cell(ws.cell(row=detail_row, column=RIGHT_COL), font=Font(bold=True, size=10, color="333333"), alignment=left_align)

row_cursor = row_header + 1
for place_name, nb_lots, fill_place in places:
    detail_row += 1
    start_r = row_cursor
    end_r = row_cursor + nb_lots - 1
    ws.cell(row=detail_row, column=RIGHT_COL, value=f"Total {place_name}")
    style_cell(ws.cell(row=detail_row, column=RIGHT_COL), font=bold_font, fill=fill_place, alignment=center)
    ws.cell(row=detail_row, column=RIGHT_COL + 1, value="")
    style_cell(ws.cell(row=detail_row, column=RIGHT_COL + 1), fill=fill_place, alignment=center)
    ws.cell(row=detail_row, column=RIGHT_END).value = f'=SUM(G{start_r}:G{end_r})'
    style_cell(ws.cell(row=detail_row, column=RIGHT_END), font=bold_font, fill=fill_place, alignment=center, number_format=currency_fmt)
    row_cursor = end_r + 1

# ============================================================
# Ajustements largeurs
# ============================================================
ws.column_dimensions['A'].width = 5
ws.column_dimensions['B'].width = 28
ws.column_dimensions['C'].width = 16
ws.column_dimensions['D'].width = 3   # séparateur
ws.column_dimensions['E'].width = 22
ws.column_dimensions['F'].width = 28
ws.column_dimensions['G'].width = 16

# Figer les volets
ws.freeze_panes = 'A5'

# ============================================================
# Sauvegarde
# ============================================================
output = "/home/user/planning-urbansoccer/budget_lots_padel.xlsx"
wb.save(output)
print(f"Fichier généré : {output}")
print(f"  - Catalogue : {len(lots)} lots")
print(f"  - Attribution : 4 lots 1er + 3 lots 2ème + 2 lots 3ème = 9 sélecteurs")

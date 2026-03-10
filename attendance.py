#!/usr/bin/env python3
"""
Module de parsing et écriture du fichier Excel de présences PSG Academy.

Structure attendue par onglet (chaque onglet = un créneau) :
  - Ligne 2 : Titre du créneau (ex: "SAMEDI 11H15")
  - Ligne 5 : Dates des séances (format DD/MM/YY)
  - Ligne 6 : Numéros de séances (S1, S2, ..., V pour vacances)
  - Lignes 7+ : Données enfants par groupe
      Col B : numéro dans le groupe
      Col C : Nom complet (NOM Prénom)
      Col D : Catégorie (U10, U11, U12, U13, BABY)
      Col E+ : 1 (présent) ou 0 (absent)
  - Lignes TOTAL : sommes par séance

Gère les groupes multiples par onglet, séparés par des lignes TOTAL/vides.
"""

import os
import re
import glob as glob_mod
from datetime import datetime
from copy import copy

import openpyxl

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def find_attendance_file():
    """Trouve le fichier Excel de présences PSG Academy."""
    patterns = [
        os.path.join(SCRIPT_DIR, "*PSG*résence*xlsx"),
        os.path.join(SCRIPT_DIR, "*PSG*resence*xlsx"),
        os.path.join(SCRIPT_DIR, "*PSG*Presence*xlsx"),
        os.path.join(SCRIPT_DIR, "*presence*PSG*xlsx"),
        os.path.join(SCRIPT_DIR, "*Présence*PSG*xlsx"),
        os.path.join(SCRIPT_DIR, "*PSG*Academy*xlsx"),
        os.path.join(SCRIPT_DIR, "PSG*.xlsx"),
    ]
    for pat in patterns:
        matches = glob_mod.glob(pat)
        if matches:
            return matches[0]
    # Fallback : chercher tout fichier avec "Présence" ou "Presence"
    for pat in [
        os.path.join(SCRIPT_DIR, "*résence*xlsx"),
        os.path.join(SCRIPT_DIR, "*resence*xlsx"),
        os.path.join(SCRIPT_DIR, "*Presence*xlsx"),
    ]:
        matches = glob_mod.glob(pat)
        # Exclure les fichiers de planning staff
        matches = [m for m in matches if "Plannings 2026" not in m]
        if matches:
            return matches[0]
    return None


def parse_attendance(filepath=None):
    """
    Parse le fichier Excel de présences.

    Retourne un dict :
    {
        "file": str,  # chemin du fichier
        "creneaux": [
            {
                "sheet_name": str,       # nom de l'onglet
                "title": str,            # titre affiché (ex: "SAMEDI 11H15")
                "slug": str,             # slug pour l'URL
                "sessions": [            # liste des colonnes séance
                    {"col": int, "label": str, "date": str|None, "is_vacation": bool}
                ],
                "groups": [
                    {
                        "name": str,     # "Groupe 1", "Groupe 2", ...
                        "kids": [
                            {
                                "row": int,       # ligne Excel (1-indexed)
                                "num": int,       # numéro dans le groupe
                                "name": str,      # nom complet
                                "category": str,  # U10, U11, etc.
                                "attendance": {    # "S1": 1, "S2": 0, ...
                                    str: int|None
                                }
                            }
                        ],
                        "total_row": int|None
                    }
                ]
            }
        ]
    }
    """
    if filepath is None:
        filepath = find_attendance_file()
    if not filepath or not os.path.exists(filepath):
        return {"file": None, "creneaux": []}

    wb = openpyxl.load_workbook(filepath, data_only=True)
    result = {"file": filepath, "creneaux": []}

    for sheet_name in wb.sheetnames:
        # Filtrer les onglets de présence
        if not _is_presence_sheet(sheet_name):
            continue

        ws = wb[sheet_name]
        creneau = _parse_sheet(ws, sheet_name)
        if creneau and creneau["groups"]:
            result["creneaux"].append(creneau)

    wb.close()
    return result


def _is_presence_sheet(name):
    """Détermine si un onglet est un onglet de présence."""
    name_lower = name.lower()
    return ("présence" in name_lower or "presence" in name_lower)


def _slugify(text):
    """Crée un slug URL-friendly."""
    import unicodedata
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^\w\s-]", "", text.lower())
    return re.sub(r"[-\s]+", "-", text).strip("-")


def _parse_sheet(ws, sheet_name):
    """Parse un onglet de présence."""
    max_row = ws.max_row or 50
    max_col = ws.max_column or 50

    # 1. Trouver le titre (ligne 2 ou première cellule non vide fusionnée)
    title = ""
    for row in range(1, min(4, max_row + 1)):
        for col in range(1, min(10, max_col + 1)):
            val = ws.cell(row=row, column=col).value
            if val and isinstance(val, str) and len(val) > 3:
                title = val.strip()
                break
        if title:
            break
    if not title:
        # Extraire du nom de l'onglet
        title = sheet_name.replace(" Présence", "").replace(" Presence", "").strip()

    slug = _slugify(title or sheet_name)

    # 2. Trouver la ligne des sessions (contient S1, S2, etc.)
    session_row = None
    date_row = None
    for row in range(1, min(10, max_row + 1)):
        row_vals = []
        for col in range(1, min(max_col + 1, 100)):
            v = ws.cell(row=row, column=col).value
            if v is not None:
                row_vals.append((col, str(v).strip()))
        # Chercher S1 dans cette ligne
        for col, v in row_vals:
            if v == "S1":
                session_row = row
                # La ligne des dates est juste au-dessus
                date_row = row - 1 if row > 1 else None
                break
        if session_row:
            break

    if not session_row:
        return None

    # 3. Parser les sessions
    sessions = []
    first_session_col = None
    for col in range(1, min(max_col + 1, 100)):
        val = ws.cell(row=session_row, column=col).value
        if val is None:
            continue
        label = str(val).strip()
        if re.match(r"^S\d+$", label) or label == "V" or label == "F":
            is_vacation = label in ("V", "F")
            # Récupérer la date correspondante
            date_str = None
            if date_row:
                date_val = ws.cell(row=date_row, column=col).value
                if date_val:
                    if isinstance(date_val, datetime):
                        date_str = date_val.strftime("%d/%m/%y")
                    else:
                        date_str = str(date_val).strip()

            sessions.append({
                "col": col,
                "label": label,
                "date": date_str,
                "is_vacation": is_vacation,
            })
            if first_session_col is None and not is_vacation:
                first_session_col = col

    if not sessions:
        return None

    # 4. Parser les enfants et groupes
    groups = []
    current_group = {"name": "Groupe 1", "kids": [], "total_row": None}
    group_num = 1
    data_start_row = session_row + 1

    for row in range(data_start_row, max_row + 1):
        col_b = ws.cell(row=row, column=2).value  # numéro
        col_c = ws.cell(row=row, column=3).value  # nom
        col_d = ws.cell(row=row, column=4).value  # catégorie

        # Détecter ligne TOTAL
        if col_c and isinstance(col_c, str) and "TOTAL" in col_c.upper():
            current_group["total_row"] = row
            if current_group["kids"]:
                groups.append(current_group)
            group_num += 1
            current_group = {"name": f"Groupe {group_num}", "kids": [], "total_row": None}
            continue

        # Détecter un enfant : col_b est un numéro et col_c est un nom
        if col_b is not None and col_c and isinstance(col_c, str) and col_c.strip():
            try:
                num = int(col_b)
            except (ValueError, TypeError):
                continue

            name = col_c.strip()
            category = str(col_d).strip() if col_d else ""

            # Lire les présences
            attendance = {}
            for sess in sessions:
                cell_val = ws.cell(row=row, column=sess["col"]).value
                if cell_val is not None:
                    try:
                        attendance[sess["label"]] = int(cell_val)
                    except (ValueError, TypeError):
                        attendance[sess["label"]] = None
                else:
                    attendance[sess["label"]] = None

            current_group["kids"].append({
                "row": row,
                "num": num,
                "name": name,
                "category": category,
                "attendance": attendance,
            })

    # Ajouter le dernier groupe s'il a des enfants
    if current_group["kids"]:
        groups.append(current_group)

    return {
        "sheet_name": sheet_name,
        "title": title,
        "slug": slug,
        "sessions": sessions,
        "groups": groups,
    }


def save_attendance(filepath, sheet_name, updates):
    """
    Écrit les présences dans le fichier Excel.

    updates : liste de {"row": int, "col": int, "value": int}
    """
    wb = openpyxl.load_workbook(filepath)
    ws = wb[sheet_name]

    for u in updates:
        cell = ws.cell(row=u["row"], column=u["col"])
        cell.value = u["value"]

    # Recalculer les TOTAL si possible
    _update_totals(ws)

    wb.save(filepath)
    wb.close()


def _update_totals(ws):
    """Recalcule les lignes TOTAL après mise à jour."""
    max_row = ws.max_row or 50
    max_col = ws.max_column or 50

    # Trouver la ligne session pour identifier les colonnes de données
    session_row = None
    for row in range(1, min(10, max_row + 1)):
        for col in range(1, min(max_col + 1, 100)):
            if str(ws.cell(row=row, column=col).value or "").strip() == "S1":
                session_row = row
                break
        if session_row:
            break

    if not session_row:
        return

    # Identifier les colonnes de sessions
    session_cols = []
    for col in range(1, min(max_col + 1, 100)):
        val = str(ws.cell(row=session_row, column=col).value or "").strip()
        if re.match(r"^S\d+$", val) or val in ("V", "F"):
            session_cols.append(col)

    # Trouver les lignes TOTAL et recalculer
    data_start = session_row + 1
    group_start = data_start

    for row in range(data_start, max_row + 1):
        col_c = ws.cell(row=row, column=3).value
        if col_c and isinstance(col_c, str) and "TOTAL" in col_c.upper():
            # Calculer somme pour chaque colonne session
            for col in session_cols:
                total = 0
                for r in range(group_start, row):
                    val = ws.cell(row=r, column=col).value
                    if val is not None:
                        try:
                            total += int(val)
                        except (ValueError, TypeError):
                            pass
                ws.cell(row=row, column=col).value = total
            group_start = row + 1


def get_current_session(sessions):
    """
    Détermine la séance en cours ou la prochaine séance à remplir.
    Retourne le label de la séance (ex: "S12").
    """
    today = datetime.now()

    # Chercher la dernière séance non-vacances avec une date <= aujourd'hui
    last_session = None
    for sess in sessions:
        if sess["is_vacation"]:
            continue
        if sess["date"]:
            try:
                # Essayer différents formats de date
                for fmt in ("%d/%m/%y", "%d/%m/%Y"):
                    try:
                        d = datetime.strptime(sess["date"], fmt)
                        if d.date() <= today.date():
                            last_session = sess
                        break
                    except ValueError:
                        continue
            except Exception:
                continue

    if last_session:
        return last_session["label"]

    # Fallback : première séance non-vacances
    for sess in sessions:
        if not sess["is_vacation"]:
            return sess["label"]
    return sessions[0]["label"] if sessions else None


def get_session_col(sessions, label):
    """Retourne le numéro de colonne pour un label de séance."""
    for sess in sessions:
        if sess["label"] == label:
            return sess["col"]
    return None

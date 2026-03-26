#!/usr/bin/env python3
"""
Synchronise les présences JSON (GitHub) → Excel.

Lit les fichiers data/presences-*.json (modifiés via l'interface web)
et met à jour le fichier Excel de présences correspondant.

Usage :
    python sync_presences.py
"""

import json
import os
import glob

from attendance import find_attendance_file, parse_attendance, save_attendance, get_session_col

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def sync():
    """Synchronise tous les JSON de présences vers l'Excel."""
    filepath = find_attendance_file()
    if not filepath:
        print("Aucun fichier Excel de présences trouvé.")
        return

    print(f"Fichier Excel : {filepath}")

    # Parse current Excel to get structure
    data = parse_attendance(filepath)
    if not data["creneaux"]:
        print("Aucun créneau trouvé dans l'Excel.")
        return

    # Build lookup: slug -> creneau data
    creneau_map = {c["slug"]: c for c in data["creneaux"]}

    # Find all JSON files
    json_files = sorted(glob.glob(os.path.join(SCRIPT_DIR, "data", "presences-*.json")))
    if not json_files:
        print("Aucun fichier JSON de présences trouvé dans data/.")
        return

    total_updates = 0

    for jf in json_files:
        with open(jf, "r", encoding="utf-8") as f:
            json_data = json.load(f)

        slug = json_data.get("slug", "")
        sheet_name = json_data.get("sheet_name", "")
        title = json_data.get("title", slug)

        if slug not in creneau_map:
            print(f"  ⚠ Créneau '{title}' ({slug}) non trouvé dans l'Excel — ignoré")
            continue

        excel_creneau = creneau_map[slug]
        sessions = excel_creneau["sessions"]

        # Compare JSON vs Excel and collect updates
        updates = []
        for g_idx, json_group in enumerate(json_data.get("groups", [])):
            for json_kid in json_group.get("kids", []):
                row = json_kid["row"]
                for sess_label, json_val in json_kid.get("attendance", {}).items():
                    if json_val is None:
                        continue
                    col = get_session_col(sessions, sess_label)
                    if col is None:
                        continue
                    # Find Excel value
                    excel_val = _find_excel_value(excel_creneau, row, sess_label)
                    if excel_val != json_val:
                        updates.append({"row": row, "col": col, "value": json_val})

        if updates:
            save_attendance(filepath, sheet_name, updates)
            print(f"  ✓ {title} : {len(updates)} cellules mises à jour")
            total_updates += len(updates)
        else:
            print(f"  — {title} : aucun changement")

    print(f"\nTotal : {total_updates} cellules mises à jour dans {filepath}")


def _find_excel_value(creneau, row, sess_label):
    """Trouve la valeur Excel pour un enfant/session donnés."""
    for group in creneau["groups"]:
        for kid in group["kids"]:
            if kid["row"] == row:
                return kid["attendance"].get(sess_label)
    return None


if __name__ == "__main__":
    sync()

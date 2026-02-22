#!/usr/bin/env python3
"""
Générateur de plannings Urban 7D — Abonnement calendrier.

Convertit les fichiers Excel « Plannings YYYY SXX.xlsx » en :
  - fichiers ICS (un par employé, cumulatif toutes semaines)
  - pages HTML (une par semaine, avec liens webcal://)
  - fichiers JSON (métadonnées par semaine)

Usage :
    python generate.py

Architecture :
    Excel (source) ──► generate.py ──► ics/ + HTML + data/
    Héberger sur un serveur web (GitHub Pages, Netlify, etc.)
    Les employés s'abonnent via webcal://domaine/ics/nom.ics

Le script détecte automatiquement tous les fichiers
« Plannings YYYY SXX.xlsx » présents dans le répertoire courant.
"""

import openpyxl
import json
import os
import re
from datetime import datetime, timedelta

# ── Mapping codes → noms lisibles + couleurs néon (basées sur l'Excel) ────

CODE_NAMES = {
    "VDC":   "Vie de centre",
    "L-REG": "Régisseur League",
    "CUP-L": "Cup League",
    "CUP-R": "Cup Régisseur",
    "STAGE": "Stage",
    "STA-E": "Stage encadrement",
    "C-PAD": "Cours Padel",
    "PAD-A": "Padel animation",
    "ANNIV": "Anniversaire",
    "INVEN": "Inventaire",
    "MAL":   "Maladie",
    "REU":   "Réunion",
    "EV-RE": "Événement régisseur",
    "AIDE":  "Aide",
    "P25M":  "P25M",
    "PSG":   "PSG Academy",
}

# Couleurs néon par code — inspirées de l'Excel avec effet glow
CODE_COLORS = {
    "VDC":   {"bg": "rgba(255,255,255,0.12)", "border": "#ffffff",  "text": "#ffffff"},
    "L-REG": {"bg": "rgba(180,100,255,0.20)", "border": "#b464ff",  "text": "#d4a0ff"},
    "CUP-L": {"bg": "rgba(255,255,255,0.12)", "border": "#ffffff",  "text": "#ffffff"},
    "CUP-R": {"bg": "rgba(100,220,60,0.20)",  "border": "#64dc3c",  "text": "#90ff70"},
    "STAGE": {"bg": "rgba(100,230,255,0.15)", "border": "#64e6ff",  "text": "#a0f0ff"},
    "STA-E": {"bg": "rgba(100,230,255,0.15)", "border": "#64e6ff",  "text": "#a0f0ff"},
    "C-PAD": {"bg": "rgba(180,180,180,0.15)", "border": "#b4b4b4",  "text": "#d0d0d0"},
    "PAD-A": {"bg": "rgba(180,180,180,0.15)", "border": "#b4b4b4",  "text": "#d0d0d0"},
    "ANNIV": {"bg": "rgba(0,176,240,0.25)",   "border": "#00b0f0",  "text": "#60d0ff"},
    "INVEN": {"bg": "rgba(255,192,0,0.25)",   "border": "#ffc000",  "text": "#ffd060"},
    "MAL":   {"bg": "rgba(255,80,80,0.20)",   "border": "#ff5050",  "text": "#ff8080"},
    "REU":   {"bg": "rgba(255,192,0,0.25)",   "border": "#ffc000",  "text": "#ffd060"},
    "EV-RE": {"bg": "rgba(100,220,60,0.20)",  "border": "#64dc3c",  "text": "#90ff70"},
    "AIDE":  {"bg": "rgba(0,176,240,0.25)",   "border": "#00b0f0",  "text": "#60d0ff"},
    "P25M":  {"bg": "rgba(255,120,50,0.25)",  "border": "#ff7832",  "text": "#ff9850"},
    "PSG":   {"bg": "rgba(255,120,50,0.25)",  "border": "#ff7832",  "text": "#ff9850"},
}
DEFAULT_COLOR = {"bg": "rgba(255,255,255,0.10)", "border": "#888888", "text": "#cccccc"}

COLS = ["B", "C", "D", "E", "F", "G", "H"]

FRENCH_MONTHS = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre",
}

FRENCH_MONTHS_LOWER = {
    1: "janvier", 2: "février", 3: "mars", 4: "avril",
    5: "mai", 6: "juin", 7: "juillet", 8: "août",
    9: "septembre", 10: "octobre", 11: "novembre", 12: "décembre",
}

# ── Utilitaires ────────────────────────────────────────────────────────────


def format_note_date_fr(date_str):
    """Formate '2026-02-20' ou '2026-02-20T14:30' en français."""
    if not date_str:
        return ""
    if "T" in date_str:
        date_part, time_part = date_str.split("T", 1)
        y, m, d = date_part.split("-")
        h, mi = time_part.split(":")[:2]
        return f"{int(d)} {FRENCH_MONTHS_LOWER[int(m)]} {y} à {int(h)}h{mi}"
    y, m, d = date_str.split("-")
    return f"{int(d)} {FRENCH_MONTHS_LOWER[int(m)]} {y}"


def slug(name):
    """BONILLO Matthieu -> bonillo-matthieu"""
    s = name.lower()
    for old, new in [("ï", "i"), ("é", "e"), ("è", "e"), ("ê", "e"),
                     ("ô", "o"), ("ü", "u"), ("ù", "u"), ("û", "u"),
                     ("à", "a"), ("â", "a"), ("ç", "c")]:
        s = s.replace(old, new)
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")


def week_dates(year, week):
    """Calcule les dates Lundi→Dimanche à partir de l'année/semaine ISO."""
    monday = datetime.fromisocalendar(year, week, 1)
    return {col: monday + timedelta(days=i) for i, col in enumerate(COLS)}


def format_date_range(year, week):
    """Retourne « 2 → 8 Mars » ou « 28 Février → 6 Mars »."""
    monday = datetime.fromisocalendar(year, week, 1)
    sunday = monday + timedelta(days=6)
    if monday.month == sunday.month:
        return f"{monday.day} \u2192 {sunday.day} {FRENCH_MONTHS[monday.month]}"
    return (f"{monday.day} {FRENCH_MONTHS[monday.month]} \u2192 "
            f"{sunday.day} {FRENCH_MONTHS[sunday.month]}")


def discover_excel_files(directory="."):
    """Trouve tous les fichiers « Plannings YYYY SXX.xlsx »."""
    pattern = re.compile(r"Plannings\s+(\d{4})\s+S(\d+)(?:\s+v\d+)?\.xlsx", re.IGNORECASE)
    files = []
    for f in sorted(os.listdir(directory)):
        m = pattern.match(f)
        if m:
            files.append({
                "filename": os.path.join(directory, f) if directory != "." else f,
                "year": int(m.group(1)),
                "week": int(m.group(2)),
            })
    files.sort(key=lambda x: (x["year"], x["week"]))
    return files


# ── Parsing Excel ──────────────────────────────────────────────────────────


def get_cell(ws, row, col):
    return ws[f"{col}{row}"].value


def normalize_time_str(val):
    """Normalise une valeur de cellule horaire en chaîne « HH:MM/HH:MM[+] ».

    Gère :
    - str  « 08:00/10:00 », « 8:00/10:00 », « 19:00/00:30+ »
    - datetime (Excel formate la cellule en Heure) → converti en str
    - Retourne None si non reconnu.
    """
    if isinstance(val, datetime):
        # Excel time-formatted cell: openpyxl returns datetime(1900,1,1,H,M,S)
        # On ne peut extraire qu'une seule heure, pas un intervalle → warning
        return None
    if not isinstance(val, str):
        return None
    s = val.strip()
    if not s:
        return None
    # Accepter les heures à 1 ou 2 chiffres : « 8:00/10:00 » → « 08:00/10:00 »
    m = re.match(r"^(\d{1,2}):(\d{2})/(\d{1,2}):(\d{2})(\+?)$", s)
    if not m:
        return None
    return f"{int(m.group(1)):02d}:{m.group(2)}/{int(m.group(3)):02d}:{m.group(4)}{m.group(5)}"


def parse_time(time_str, base_date):
    """Parse « 08:00/10:00 » ou « 19:00/00:30+ » en (start_dt, end_dt).

    Le suffixe « + » indique explicitement que l'heure de fin est le
    lendemain (ex : CUP-R 19:00/00:30+  →  19h → 0h30 le jour suivant).
    Même sans « + », si end ≤ start le lendemain est détecté automatiquement.
    """
    parts = time_str.strip().split("/")
    if len(parts) != 2:
        return None

    start_str = parts[0].strip()
    end_str = parts[1].strip()

    next_day = end_str.endswith("+")
    if next_day:
        end_str = end_str.rstrip("+")

    sh, sm = int(start_str.split(":")[0]), int(start_str.split(":")[1])
    eh, em = int(end_str.split(":")[0]), int(end_str.split(":")[1])

    # Gérer 24:00 comme minuit du jour suivant
    start_extra = 0
    if sh >= 24:
        sh -= 24
        start_extra = 1
    end_extra = 0
    if eh >= 24:
        eh -= 24
        end_extra = 1

    start_dt = base_date.replace(hour=sh, minute=sm, second=0) + timedelta(days=start_extra)
    end_dt = base_date.replace(hour=eh, minute=em, second=0) + timedelta(days=end_extra)

    # « + » explicite OU détection automatique si fin ≤ début
    if next_day or (end_dt <= start_dt):
        end_dt += timedelta(days=1)

    return (start_dt, end_dt)


def parse_employees(ws, dates, week_num):
    """Parse tous les employés et leurs créneaux depuis la feuille Planning."""
    employees = {}
    current_name = None
    current_rows = []

    for row in range(5, ws.max_row + 1):
        name_cell = get_cell(ws, row, "A")
        if name_cell and isinstance(name_cell, str) and name_cell.strip():
            if current_name:
                employees[current_name] = parse_shifts(ws, current_rows, dates, week_num, current_name)
            current_name = name_cell.strip()
            current_rows = [row]
        elif current_name:
            current_rows.append(row)

    if current_name:
        employees[current_name] = parse_shifts(ws, current_rows, dates, week_num, current_name)

    return employees


def parse_shifts(ws, rows, dates, week_num, employee_name=""):
    """Parse les créneaux d'un employé à partir de ses lignes."""
    events = []
    warnings = []

    i = 0
    while i < len(rows):
        row = rows[i]
        has_codes = False
        codes = {}
        for col in COLS:
            val = get_cell(ws, row, col)
            if val and isinstance(val, str):
                val = val.strip()
                if val and not re.match(r"^\d{1,2}:\d{2}/\d{1,2}:\d{2}", val):
                    has_codes = True
                    codes[col] = val

        if has_codes:
            times = {}
            if i + 1 < len(rows):
                time_row = rows[i + 1]
                for col in COLS:
                    raw_val = get_cell(ws, time_row, col)
                    if raw_val is None:
                        continue
                    normalized = normalize_time_str(raw_val)
                    if normalized:
                        times[col] = normalized
                    elif isinstance(raw_val, datetime):
                        warnings.append(
                            f"  /!\\ {employee_name} ligne {time_row} col {col} : "
                            f"cellule format\u00e9e en Heure ({raw_val.strftime('%H:%M')}), "
                            f"convertir en texte dans Excel"
                        )
            else:
                for col, code in codes.items():
                    if col in dates:
                        warnings.append(
                            f"  /!\\ {employee_name} ligne {row} col {col} : "
                            f"code \u00ab {code} \u00bb sans ligne horaire en dessous"
                        )

            for col, code in codes.items():
                if col in dates:
                    if col in times:
                        parsed = parse_time(times[col], dates[col])
                        if parsed:
                            label = CODE_NAMES.get(code, code)
                            events.append({
                                "code": code,
                                "label": label,
                                "start": parsed[0],
                                "end": parsed[1],
                                "week": week_num,
                            })
                    else:
                        warnings.append(
                            f"  /!\\ {employee_name} ligne {row} col {col} : "
                            f"code \u00ab {code} \u00bb sans horaire trouv\u00e9"
                        )

            i += 2
        else:
            i += 1

    for w in warnings:
        print(w)

    events.sort(key=lambda e: e["start"])
    return events


# ── Génération ICS (abonnement calendrier) ─────────────────────────────────


def generate_ics(name, events, week_notes=None):
    """Génère le contenu ICS pour un employé (toutes semaines confondues).

    Chaque fichier ICS contient TOUS les événements de l'employé, ce qui
    permet à l'abonnement calendrier de rester à jour automatiquement.
    week_notes: dict {week_num: notes_data} pour ajouter les commentaires.
    """
    if week_notes is None:
        week_notes = {}
    s = slug(name)
    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Planning Urban 7D//FR",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:Planning {name}",
        "X-WR-TIMEZONE:Europe/Paris",
        # Intervalle de rafraîchissement pour les clients calendrier
        "REFRESH-INTERVAL;VALUE=DURATION:PT12H",
        "X-PUBLISHED-TTL:PT12H",
        "BEGIN:VTIMEZONE",
        "TZID:Europe/Paris",
        "BEGIN:STANDARD",
        "TZOFFSETFROM:+0200",
        "TZOFFSETTO:+0100",
        "TZNAME:CET",
        "DTSTART:19701025T030000",
        "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU",
        "END:STANDARD",
        "BEGIN:DAYLIGHT",
        "TZOFFSETFROM:+0100",
        "TZOFFSETTO:+0200",
        "TZNAME:CEST",
        "DTSTART:19700329T020000",
        "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU",
        "END:DAYLIGHT",
        "END:VTIMEZONE",
    ]

    # Grouper par semaine pour des UIDs stables
    by_week = {}
    for evt in events:
        w = evt["week"]
        if w not in by_week:
            by_week[w] = []
        by_week[w].append(evt)

    for week_num in sorted(by_week.keys()):
        # Build description with weekly notes if available
        wn = week_notes.get(week_num, {})
        week_comment = wn.get("comment", "")
        week_updates = wn.get("updates", [])
        extra_desc = ""
        if week_comment:
            extra_desc += "\\n---\\n" + week_comment
        for upd in week_updates:
            upd_text = upd.get("text", "")
            upd_date = upd.get("date", "")
            if upd_text:
                date_label = format_note_date_fr(upd_date) if upd_date else ""
                prefix = f"MAJ {date_label}: " if date_label else "MAJ: "
                extra_desc += "\\n" + prefix + upd_text

        for i, evt in enumerate(by_week[week_num], 1):
            dt_start = evt["start"].strftime("%Y%m%dT%H%M%S")
            dt_end = evt["end"].strftime("%Y%m%dT%H%M%S")
            # Escape for ICS DESCRIPTION
            desc = evt['label']
            if extra_desc:
                desc += extra_desc
            desc_escaped = desc.replace("\\", "\\\\").replace("\n", "\\n").replace(",", "\\,").replace(";", "\\;")
            lines.extend([
                "BEGIN:VEVENT",
                f"UID:{s}-s{week_num}-{i}@urban7d",
                f"DTSTAMP:{dt_start}",
                f"DTSTART;TZID=Europe/Paris:{dt_start}",
                f"DTEND;TZID=Europe/Paris:{dt_end}",
                f"SUMMARY:{evt['label']}",
                f"DESCRIPTION:{desc_escaped}",
                "END:VEVENT",
            ])

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


# ── Génération HTML ────────────────────────────────────────────────────────


def build_events_json(week_employees):
    """Construit les données JSON des événements pour injection dans le HTML."""
    data = {}
    for name, evts in week_employees.items():
        data[name] = {
            "slug": slug(name),
            "events": [{
                "code": e["code"],
                "label": e["label"],
                "start": e["start"].strftime("%Y-%m-%dT%H:%M"),
                "end": e["end"].strftime("%Y-%m-%dT%H:%M"),
                "day": e["start"].weekday(),
            } for e in evts],
        }
    return json.dumps(data, ensure_ascii=False)


def load_week_notes(week_num):
    """Charge les notes de semaine depuis notes/SXX.json."""
    path = f"notes/S{week_num}.json"
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, IOError):
            pass
    return {"comment": "", "updates": []}


def generate_html(week_employees, week_num, year, all_weeks):
    """Génère la page HTML avec preview timeline + vue individuelle + abonnement."""
    date_range = format_date_range(year, week_num)
    events_json = build_events_json(week_employees)
    colors_json = json.dumps(CODE_COLORS, ensure_ascii=False)
    default_color_json = json.dumps(DEFAULT_COLOR, ensure_ascii=False)
    notes_data = load_week_notes(week_num)
    notes_json = json.dumps(notes_data, ensure_ascii=False)

    DAYS_SHORT = ["Lun", "Mar", "Mer", "Jeu", "Ven", "Sam", "Dim"]
    DAYS_FULL = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
    monday = datetime.fromisocalendar(year, week_num, 1)
    day_labels_json = json.dumps([
        f"{DAYS_SHORT[i]} {(monday + timedelta(days=i)).day:02d}/{(monday + timedelta(days=i)).month:02d}"
        for i in range(7)
    ], ensure_ascii=False)
    day_labels_full_json = json.dumps([
        f"{DAYS_FULL[i]} {(monday + timedelta(days=i)).day:02d}/{(monday + timedelta(days=i)).month:02d}"
        for i in range(7)
    ], ensure_ascii=False)

    week_tabs = ""
    for w in sorted(all_weeks):
        cls = ' active' if w == week_num else ''
        href = '#' if w == week_num else f'S{w}.html'
        week_tabs += f'            <a href="{href}" class="week-tab{cls}">S{w}</a>\n'

    employee_buttons = ""
    for name in week_employees:
        s = slug(name)
        has_events = len(week_employees[name]) > 0
        if has_events:
            employee_buttons += (
                f'            <button class="employee-btn" data-name="{name}" '
                f'data-slug="{s}">{name}</button>\n'
            )
        else:
            employee_buttons += (
                f'            <div class="employee-btn repos">{name} '
                f'<span class="badge">Repos</span></div>\n'
            )

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Planning Urban 7D - S{week_num}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=Tahoma&display=swap" rel="stylesheet">
    <style>
        @font-face {{ font-family: 'Heading'; src: local('GT Pressura Mono Bold'), local('Space Mono'); }}
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: Tahoma, 'Inter', sans-serif;
            background: #1E1E1E;
            min-height: 100vh;
            padding: 15px;
            color: #fff;
            position: relative;
        }}
        body::before {{
            content: '';
            position: fixed;
            inset: 0;
            z-index: 0;
            pointer-events: none;
            background: url('bg-team.jpg') center center / cover no-repeat;
            opacity: 0.13;
        }}
        .container {{ position: relative; z-index: 1; max-width: 600px; margin: 0 auto; }}

        /* ── Header ── */
        .header {{ text-align: center; margin-bottom: 12px; padding: 10px 10px 8px; }}
        h1 {{ font-family: 'Space Mono', 'GT Pressura Mono Bold', monospace;
              color: #FF7832; font-size: 18px; font-weight: 700; margin-bottom: 2px;
              text-transform: uppercase; letter-spacing: 1px;
              text-shadow: 0 0 30px rgba(255,120,50,0.3); }}
        .subtitle {{ color: #888; font-size: 12px; }}
        .dates {{ color: #FF7832; font-size: 14px; font-weight: 600;
                  background: rgba(255,120,50,0.1); padding: 6px 14px;
                  border-radius: 20px; display: inline-block; margin-top: 6px;
                  border: 1px solid rgba(255,120,50,0.2); }}

        /* ── Week selector ── */
        .week-selector {{ display: flex; justify-content: center; gap: 6px;
                          margin-bottom: 15px; flex-wrap: wrap; }}
        .week-tab {{ padding: 8px 14px; background: rgba(255,255,255,0.04);
                     border: 1px solid rgba(255,255,255,0.08); border-radius: 20px;
                     color: #666; text-decoration: none; font-weight: 500; font-size: 13px;
                     transition: all 0.2s; }}
        .week-tab:hover {{ background: rgba(255,120,50,0.1); border-color: rgba(255,120,50,0.3); color: #FF7832; }}
        .week-tab.active {{ background: #FF7832; border-color: #FF7832; color: white;
                            box-shadow: 0 0 15px rgba(255,120,50,0.4); }}

        /* ── View toggle ── */
        .view-toggle {{ display: flex; justify-content: center; gap: 4px; margin-bottom: 15px;
                        background: rgba(255,255,255,0.04); border-radius: 12px; padding: 4px; }}
        .view-btn {{ flex: 1; padding: 8px; border: none; background: transparent;
                     color: #666; font-size: 12px; font-weight: 600; cursor: pointer;
                     border-radius: 10px; transition: all 0.2s; font-family: inherit; }}
        .view-btn.active {{ background: rgba(255,120,50,0.15); color: #FF7832;
                            box-shadow: 0 0 10px rgba(255,120,50,0.2); }}

        /* ── Day tabs ── */
        .day-tabs {{ display: flex; gap: 3px; margin-bottom: 12px; overflow-x: auto;
                     padding-bottom: 4px; -webkit-overflow-scrolling: touch;
                     scrollbar-width: none; }}
        .day-tabs::-webkit-scrollbar {{ display: none; }}
        .day-tab {{ padding: 6px 8px; background: rgba(255,255,255,0.04);
                    border: 1px solid rgba(255,255,255,0.08); border-radius: 8px;
                    color: #666; font-size: 10px; font-weight: 600; cursor: pointer;
                    white-space: nowrap; transition: all 0.2s; flex: 1; min-width: 0;
                    text-align: center; }}
        .day-tab.active {{ background: rgba(255,120,50,0.15); border-color: rgba(255,120,50,0.3);
                           color: #FF7832; }}

        /* ── Timeline (vue Journée) ── */
        .timeline {{ position: relative; margin-bottom: 20px;
                     overflow-x: auto; -webkit-overflow-scrolling: touch; }}
        /* ── Scrollbar orange néon ── */
        .timeline::-webkit-scrollbar {{ height: 6px; }}
        .timeline::-webkit-scrollbar-track {{ background: rgba(255,255,255,0.04); border-radius: 3px; }}
        .timeline::-webkit-scrollbar-thumb {{ background: #FF7832; border-radius: 3px;
                                              box-shadow: 0 0 8px rgba(255,120,50,0.6); }}
        .timeline {{ scrollbar-width: thin; scrollbar-color: #FF7832 rgba(255,255,255,0.04); }}
        .timeline-inner {{ min-width: 500px; }}
        .time-markers {{ display: flex; justify-content: space-between; padding: 0 0 6px 0;
                         border-bottom: 1px solid rgba(255,255,255,0.06); margin-bottom: 8px; }}
        .time-marker {{ font-size: 9px; color: #555; font-weight: 500; }}
        .timeline-row {{ display: flex; align-items: center; margin-bottom: 4px; }}
        .tl-name {{ width: 70px; font-size: 10px; color: #aaa; font-weight: 500;
                    flex-shrink: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
                    padding-right: 6px; cursor: pointer; transition: color 0.2s; }}
        .tl-name:hover {{ color: #FF7832; }}
        .tl-bar-container {{ flex: 1; position: relative; height: 26px;
                             background: rgba(255,255,255,0.02); border-radius: 5px; }}
        .tl-bar {{ position: absolute; height: 100%; border-radius: 5px;
                   display: flex; align-items: center; justify-content: center;
                   font-size: 9px; font-weight: 600; overflow: hidden;
                   border-left: 2px solid; transition: all 0.2s;
                   cursor: default; }}
        .tl-bar:hover {{ filter: brightness(1.3); z-index: 2;
                         box-shadow: 0 0 12px var(--glow-color); }}
        .tl-bar .bar-label {{ padding: 0 4px; white-space: nowrap; }}

        /* ── Employee list (vue Staff) ── */
        .employee-list {{ display: flex; flex-direction: column; gap: 6px; margin-bottom: 15px; }}
        .employee-btn {{ display: flex; align-items: center; justify-content: space-between;
                         padding: 12px 14px; background: rgba(255,255,255,0.04);
                         border-radius: 10px; color: white; font-weight: 500; font-size: 13px;
                         border: 1px solid rgba(255,255,255,0.08); cursor: pointer;
                         transition: all 0.2s; font-family: inherit; width: 100%; text-align: left; }}
        .employee-btn:hover {{ background: rgba(255,120,50,0.1); border-color: rgba(255,120,50,0.3);
                               transform: translateX(4px); }}
        .employee-btn.repos {{ color: #444; cursor: default; pointer-events: none; }}
        .badge {{ font-size: 10px; padding: 3px 8px; background: rgba(255,255,255,0.06);
                  border-radius: 15px; color: #444; }}

        /* ── Individual preview (modal) ── */
        .modal-overlay {{ display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.85);
                          z-index: 100; justify-content: center; align-items: flex-start;
                          padding: 15px 10px; overflow-y: auto; }}
        .modal-overlay.open {{ display: flex; }}
        .modal {{ background: #12121e; border-radius: 14px; width: 100%; max-width: 500px;
                  border: 1px solid rgba(255,255,255,0.08); overflow: hidden; }}
        .modal-header {{ padding: 14px 16px; display: flex; justify-content: space-between;
                         align-items: center; border-bottom: 1px solid rgba(255,255,255,0.06); }}
        .modal-header h2 {{ font-size: 16px; color: #FF7832; font-weight: 700; }}
        .modal-close {{ background: none; border: none; color: #666; font-size: 24px;
                        cursor: pointer; padding: 0 5px; line-height: 1; }}
        .modal-close:hover {{ color: #fff; }}
        .modal-body {{ padding: 12px 14px; }}
        .modal-day {{ margin-bottom: 12px; }}
        .modal-day-title {{ font-size: 11px; color: #666; font-weight: 600;
                            text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 5px; }}
        .modal-event {{ display: flex; align-items: center; gap: 8px; padding: 8px 10px;
                        border-radius: 8px; margin-bottom: 3px; border-left: 3px solid; }}
        .modal-event .ev-time {{ font-size: 11px; font-weight: 600; white-space: nowrap;
                                 min-width: 80px; }}
        .modal-event .ev-label {{ font-size: 12px; font-weight: 500; }}
        .modal-footer {{ padding: 12px 16px; border-top: 1px solid rgba(255,255,255,0.06);
                         text-align: center; }}
        .subscribe-btn {{ display: inline-flex; align-items: center; gap: 8px;
                          padding: 10px 24px; background: #FF7832; color: white;
                          border: none; border-radius: 25px; font-size: 13px; font-weight: 600;
                          cursor: pointer; font-family: inherit; transition: all 0.2s;
                          text-decoration: none;
                          box-shadow: 0 0 20px rgba(255,120,50,0.3); }}
        .subscribe-btn:hover {{ background: #ff9050;
                                box-shadow: 0 0 30px rgba(255,120,50,0.5); transform: scale(1.02); }}

        .no-events {{ text-align: center; padding: 30px; color: #444; font-size: 13px; }}

        /* ── Custom checkbox (DA) ── */
        .tl-check-wrap {{ flex-shrink: 0; margin-left: 6px; }}
        .tl-check {{ display: none; }}
        .tl-check-label {{ display: block; width: 20px; height: 20px; border-radius: 5px;
                           border: 2px solid rgba(255,255,255,0.15); cursor: pointer;
                           transition: all 0.2s; position: relative; }}
        .tl-check-label:hover {{ border-color: rgba(255,120,50,0.5); }}
        .tl-check:checked + .tl-check-label {{ background: #FF7832; border-color: #FF7832;
                                                box-shadow: 0 0 8px rgba(255,120,50,0.4); }}
        .tl-check:checked + .tl-check-label::after {{ content: ''; position: absolute;
            left: 5px; top: 2px; width: 5px; height: 9px;
            border: solid #fff; border-width: 0 2px 2px 0; transform: rotate(45deg); }}
        .export-bar {{ display: none; justify-content: center; margin-top: 12px; }}
        .export-bar.visible {{ display: flex; }}
        .export-btn {{ display: inline-flex; align-items: center; gap: 6px;
                       padding: 10px 22px; background: rgba(255,120,50,0.15); color: #FF7832;
                       border: 1px solid rgba(255,120,50,0.3); border-radius: 20px;
                       font-size: 12px; font-weight: 600; cursor: pointer; font-family: inherit;
                       transition: all 0.2s; }}
        .export-btn:hover {{ background: #FF7832; color: #fff;
                             box-shadow: 0 0 20px rgba(255,120,50,0.4); }}
        .select-all-row {{ display: flex; justify-content: flex-end; align-items: center;
                           gap: 6px; margin-bottom: 4px; padding-right: 2px; }}
        .select-all-label {{ font-size: 10px; color: #555; cursor: pointer; }}

        /* ── Notes de semaine ── */
        .week-notes {{ margin-bottom: 15px; }}
        .note-card {{ background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08);
                      border-radius: 10px; padding: 12px 14px; margin-bottom: 8px; }}
        .note-card.comment {{ border-left: 3px solid #FF7832; }}
        .note-card.update {{ border-left: 3px solid #ffc000; }}
        .note-header {{ display: flex; justify-content: space-between; align-items: center;
                        margin-bottom: 6px; }}
        .note-label {{ font-size: 10px; font-weight: 700; text-transform: uppercase;
                       letter-spacing: 0.5px; }}
        .note-label.comment {{ color: #FF7832; }}
        .note-label.update {{ color: #ffc000; }}
        .note-text {{ font-size: 12px; color: #ccc; line-height: 1.6; white-space: pre-line; }}
        .note-text:empty::before {{ content: 'Cliquer pour ajouter...'; color: #444; font-style: italic; }}
        .note-text[contenteditable=true] {{ outline: none; border: 1px solid rgba(255,120,50,0.2);
                                            border-radius: 6px; padding: 8px; min-height: 40px;
                                            background: rgba(0,0,0,0.2); }}
        .note-actions {{ display: flex; gap: 6px; }}
        .note-btn {{ background: none; border: none; color: #555; font-size: 14px;
                     cursor: pointer; padding: 2px 4px; transition: color 0.2s; }}
        .note-btn:hover {{ color: #FF7832; }}
        .note-btn.del:hover {{ color: #ff5050; }}
        .add-note-btn {{ display: flex; align-items: center; justify-content: center; gap: 6px;
                         padding: 8px; background: rgba(255,255,255,0.02);
                         border: 1px dashed rgba(255,255,255,0.1); border-radius: 10px;
                         color: #444; font-size: 11px; cursor: pointer; transition: all 0.2s;
                         font-family: inherit; width: 100%; margin-bottom: 8px; }}
        .add-note-btn:hover {{ border-color: rgba(255,120,50,0.3); color: #FF7832; }}
        .publish-btn {{ display: flex; align-items: center; justify-content: center; gap: 6px;
                        padding: 10px 16px; background: #FF7832; border: none; border-radius: 10px;
                        color: #fff; font-size: 12px; font-weight: 600; cursor: pointer;
                        transition: all 0.2s; font-family: inherit; width: 100%; margin-top: 8px;
                        box-shadow: 0 0 15px rgba(255,120,50,0.3); }}
        .publish-btn:hover {{ background: #ff9050; box-shadow: 0 0 25px rgba(255,120,50,0.5); }}
        .publish-btn:disabled {{ background: #444; box-shadow: none; cursor: not-allowed; color: #888; }}
        .publish-btn.success {{ background: #64dc3c; box-shadow: 0 0 15px rgba(100,220,60,0.3); }}
        .admin-setup {{ display: flex; align-items: center; gap: 6px; margin-top: 8px; }}
        .admin-input {{ flex: 1; padding: 8px 10px; background: rgba(0,0,0,0.3);
                        border: 1px solid rgba(255,255,255,0.1); border-radius: 8px;
                        color: #ccc; font-size: 11px; font-family: inherit; outline: none; }}
        .admin-input:focus {{ border-color: rgba(255,120,50,0.4); }}
        .admin-input::placeholder {{ color: #444; }}
        .admin-save-btn {{ padding: 8px 12px; background: rgba(255,120,50,0.15);
                           border: 1px solid rgba(255,120,50,0.3); border-radius: 8px;
                           color: #FF7832; font-size: 11px; cursor: pointer; font-family: inherit;
                           white-space: nowrap; }}
        .admin-hint {{ font-size: 10px; color: #444; margin-top: 4px; }}

        /* ── Admin edit mode ── */
        .admin-toolbar {{ display: flex; align-items: center; gap: 8px; margin-bottom: 12px;
                          padding: 8px 12px; background: rgba(255,120,50,0.08);
                          border: 1px solid rgba(255,120,50,0.2); border-radius: 10px; }}
        .edit-toggle {{ padding: 6px 14px; background: rgba(255,120,50,0.15);
                        border: 1px solid rgba(255,120,50,0.3); border-radius: 8px;
                        color: #FF7832; font-size: 11px; font-weight: 600; cursor: pointer;
                        font-family: inherit; transition: all 0.2s; }}
        .edit-toggle:hover {{ background: #FF7832; color: #fff; }}
        .edit-toggle.active {{ background: #FF7832; color: #fff;
                               box-shadow: 0 0 10px rgba(255,120,50,0.4); }}
        .admin-toolbar .label {{ font-size: 11px; color: #888; }}
        .tl-bar.editable {{ cursor: pointer; }}
        .tl-bar.editable:hover {{ outline: 2px solid #FF7832; outline-offset: 1px; }}
        .edit-popup {{ position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%);
                       z-index: 200; background: #1a1a2e; border: 1px solid rgba(255,120,50,0.3);
                       border-radius: 12px; padding: 16px; min-width: 260px;
                       box-shadow: 0 10px 40px rgba(0,0,0,0.6); }}
        .edit-popup h3 {{ font-size: 13px; color: #FF7832; margin-bottom: 10px; }}
        .edit-popup .field {{ display: flex; align-items: center; gap: 8px; margin-bottom: 8px; }}
        .edit-popup .field label {{ font-size: 11px; color: #888; min-width: 45px; }}
        .edit-popup .field input {{ flex: 1; padding: 6px 8px; background: rgba(0,0,0,0.3);
                                    border: 1px solid rgba(255,255,255,0.1); border-radius: 6px;
                                    color: #fff; font-size: 12px; font-family: inherit; outline: none; }}
        .edit-popup .field input:focus {{ border-color: rgba(255,120,50,0.4); }}
        .edit-popup .actions {{ display: flex; gap: 6px; margin-top: 10px; }}
        .edit-popup .actions button {{ flex: 1; padding: 8px; border: none; border-radius: 8px;
                                       font-size: 11px; font-weight: 600; cursor: pointer;
                                       font-family: inherit; transition: all 0.2s; }}
        .edit-popup .btn-save {{ background: #FF7832; color: #fff; }}
        .edit-popup .btn-save:hover {{ background: #ff9050; }}
        .edit-popup .btn-cancel {{ background: rgba(255,255,255,0.08); color: #888; }}
        .edit-popup .btn-cancel:hover {{ background: rgba(255,255,255,0.15); color: #fff; }}
        .edit-overlay {{ position: fixed; inset: 0; z-index: 199; background: rgba(0,0,0,0.5); }}
        .edit-status {{ font-size: 10px; color: #64dc3c; margin-left: auto; }}

        /* ── Legend ── */
        .legend {{ display: flex; flex-wrap: wrap; gap: 6px; justify-content: center;
                   margin-bottom: 15px; }}
        .legend-item {{ display: flex; align-items: center; gap: 4px; font-size: 10px;
                        color: #666; padding: 3px 8px; background: rgba(255,255,255,0.03);
                        border-radius: 6px; }}
        .legend-dot {{ width: 8px; height: 8px; border-radius: 50%; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>Planning Urban 7D</h1>
            <p class="subtitle">Semaine {week_num}</p>
            <div class="dates">{date_range}</div>
        </div>

        <div class="week-selector">
{week_tabs.rstrip()}
        </div>

        <div class="week-notes" id="week-notes"></div>

        <div class="view-toggle">
            <button class="view-btn active" data-view="day">Vue Journ\u00e9e</button>
            <button class="view-btn" data-view="staff">Vue Staff</button>
        </div>

        <!-- ── Vue Journée (timeline) ── -->
        <div id="view-day">
            <div class="day-tabs" id="day-tabs"></div>
            <div class="legend" id="legend"></div>
            <div id="select-all-container"></div>
            <div class="timeline" id="timeline"></div>
            <div class="export-bar" id="export-bar">
                <button class="export-btn" id="export-btn">Exporter la s\u00e9lection</button>
            </div>
        </div>

        <!-- ── Vue Staff (liste) ── -->
        <div id="view-staff" style="display:none;">
            <div class="employee-list">
{employee_buttons.rstrip()}
            </div>
        </div>
    </div>

    <!-- ── Modal preview individuelle ── -->
    <div class="modal-overlay" id="modal">
        <div class="modal">
            <div class="modal-header">
                <h2 id="modal-name"></h2>
                <button class="modal-close" id="modal-close">&times;</button>
            </div>
            <div class="modal-body" id="modal-body"></div>
            <div class="modal-footer">
                <a class="subscribe-btn" id="modal-subscribe" href="#">
                    S'abonner au calendrier
                </a>
            </div>
        </div>
    </div>

    <script>
    (function() {{
        var NOTES_DATA = {notes_json};
        var DATA = {events_json};
        var COLORS = {colors_json};
        var DEFAULT_C = {default_color_json};
        var DAYS = {day_labels_json};
        var DAYS_FULL = {day_labels_full_json};
        var currentDay = 0;
        var currentView = 'day';
        var MOIS_FR = ['janvier','f\u00e9vrier','mars','avril','mai','juin','juillet','ao\u00fbt','septembre','octobre','novembre','d\u00e9cembre'];

        function formatDateFR(dateStr) {{
            if (!dateStr) return '';
            var parts = dateStr.split('T');
            var dp = parts[0].split('-');
            var d = parseInt(dp[2]), m = parseInt(dp[1]) - 1, y = dp[0];
            var result = d + ' ' + MOIS_FR[m] + ' ' + y;
            if (parts[1]) {{
                var tp = parts[1].split(':');
                result += ' \u00e0 ' + parseInt(tp[0]) + 'h' + tp[1];
            }}
            return result;
        }}

        function getColor(code) {{ return COLORS[code] || DEFAULT_C; }}

        // ── Day tabs ──
        var dayTabsEl = document.getElementById('day-tabs');
        DAYS.forEach(function(label, i) {{
            var btn = document.createElement('div');
            btn.className = 'day-tab' + (i === 0 ? ' active' : '');
            btn.textContent = label;
            btn.onclick = function() {{ selectDay(i); }};
            dayTabsEl.appendChild(btn);
        }});

        function selectDay(i) {{
            currentDay = i;
            dayTabsEl.querySelectorAll('.day-tab').forEach(function(t, j) {{
                t.classList.toggle('active', j === i);
            }});
            renderTimeline();
        }}

        // ── Legend ──
        function renderLegend(codes) {{
            var el = document.getElementById('legend');
            el.innerHTML = '';
            var seen = {{}};
            codes.forEach(function(code) {{
                if (seen[code]) return;
                seen[code] = true;
                var c = getColor(code);
                var item = document.createElement('div');
                item.className = 'legend-item';
                item.innerHTML = '<div class="legend-dot" style="background:' + c.border +
                    ';box-shadow:0 0 6px ' + c.border + '"></div>' +
                    (DATA._codeNames && DATA._codeNames[code] || code);
                el.appendChild(item);
            }});
        }}

        // ── Timeline rendering ──
        function renderTimeline() {{
            var tl = document.getElementById('timeline');
            tl.innerHTML = '';

            // Collect events for this day
            var dayEvents = [];
            var allCodes = [];
            Object.keys(DATA).forEach(function(name) {{
                if (name === '_codeNames') return;
                var emp = DATA[name];
                emp.events.forEach(function(ev) {{
                    if (ev.day === currentDay) {{
                        dayEvents.push({{ name: name, ev: ev }});
                        allCodes.push(ev.code);
                    }}
                }});
            }});

            if (dayEvents.length === 0) {{
                tl.innerHTML = '<div class="no-events">Aucun cr\u00e9neau ce jour</div>';
                renderLegend([]);
                return;
            }}

            renderLegend(allCodes);

            // Scrollable inner wrapper
            var inner = document.createElement('div');
            inner.className = 'timeline-inner';

            // Find time range
            var minH = 24, maxH = 0;
            dayEvents.forEach(function(d) {{
                var s = new Date(d.ev.start);
                var e = new Date(d.ev.end);
                var sh = s.getHours() + s.getMinutes()/60;
                var eh = e.getHours() + e.getMinutes()/60;
                if (eh <= sh) eh = 24;
                if (sh < minH) minH = sh;
                if (eh > maxH) maxH = eh;
            }});
            minH = Math.floor(minH);
            maxH = Math.ceil(maxH);
            if (maxH <= minH) maxH = minH + 1;
            var range = maxH - minH;

            // Set inner width: 40px per hour for comfortable reading
            var pxPerHour = 40;
            inner.style.minWidth = (70 + range * pxPerHour) + 'px';

            // Time markers
            var markerRow = document.createElement('div');
            markerRow.className = 'timeline-row';
            var markerSpacer = document.createElement('div');
            markerSpacer.className = 'tl-name';
            markerSpacer.innerHTML = '&nbsp;';
            markerRow.appendChild(markerSpacer);
            var markers = document.createElement('div');
            markers.className = 'time-markers';
            markers.style.flex = '1';
            for (var h = minH; h <= maxH; h++) {{
                var m = document.createElement('span');
                m.className = 'time-marker';
                m.textContent = h + 'h';
                markers.appendChild(m);
            }}
            markerRow.appendChild(markers);
            inner.appendChild(markerRow);

            // Group by employee
            var byName = {{}};
            var nameOrder = [];
            dayEvents.forEach(function(d) {{
                if (!byName[d.name]) {{ byName[d.name] = []; nameOrder.push(d.name); }}
                byName[d.name].push(d.ev);
            }});

            nameOrder.forEach(function(name) {{
                var row = document.createElement('div');
                row.className = 'timeline-row';

                var nameEl = document.createElement('div');
                nameEl.className = 'tl-name';
                nameEl.textContent = name.split(' ').slice(1).join(' ') || name.split(' ')[0];
                nameEl.title = name;
                nameEl.onclick = function() {{ openModal(name); }};
                row.appendChild(nameEl);

                var barContainer = document.createElement('div');
                barContainer.className = 'tl-bar-container';

                byName[name].forEach(function(ev) {{
                    var s = new Date(ev.start);
                    var e = new Date(ev.end);
                    var sh = s.getHours() + s.getMinutes()/60;
                    var eh = e.getHours() + e.getMinutes()/60;
                    if (eh <= sh) eh = 24;

                    var left = ((sh - minH) / range) * 100;
                    var width = ((eh - sh) / range) * 100;
                    if (left < 0) left = 0;
                    if (left + width > 100) width = 100 - left;

                    var c = getColor(ev.code);
                    var bar = document.createElement('div');
                    bar.className = 'tl-bar';
                    bar.style.cssText = 'left:' + left + '%;width:' + width + '%;' +
                        'background:' + c.bg + ';border-color:' + c.border + ';color:' + c.text +
                        ';--glow-color:' + c.border + ';' +
                        'box-shadow:inset 0 0 8px rgba(255,255,255,0.05), 0 0 4px ' + c.border + '40;';
                    bar.innerHTML = '<span class="bar-label">' + ev.code + '</span>';
                    bar.title = ev.label + '\\n' +
                        s.getHours().toString().padStart(2,'0') + ':' + s.getMinutes().toString().padStart(2,'0') +
                        ' - ' + e.getHours().toString().padStart(2,'0') + ':' + e.getMinutes().toString().padStart(2,'0');
                    barContainer.appendChild(bar);
                }});

                var cbWrap = document.createElement('div');
                cbWrap.className = 'tl-check-wrap';
                var cbId = 'cb-' + name.replace(/\\s+/g, '-');
                var cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.className = 'tl-check';
                cb.id = cbId;
                cb.setAttribute('data-name', name);
                cb.onchange = updateExportBar;
                var cbLabel = document.createElement('label');
                cbLabel.className = 'tl-check-label';
                cbLabel.setAttribute('for', cbId);
                cbWrap.appendChild(cb);
                cbWrap.appendChild(cbLabel);

                row.appendChild(barContainer);
                row.appendChild(cbWrap);
                inner.appendChild(row);
            }});

            tl.appendChild(inner);

            // Select-all
            var selContainer = document.getElementById('select-all-container');
            selContainer.innerHTML = '';
            if (nameOrder.length > 0) {{
                var selRow = document.createElement('div');
                selRow.className = 'select-all-row';
                var selCbWrap = document.createElement('div');
                selCbWrap.className = 'tl-check-wrap';
                var selCb = document.createElement('input');
                selCb.type = 'checkbox';
                selCb.className = 'tl-check';
                selCb.id = 'select-all-cb';
                selCb.onchange = function() {{
                    document.querySelectorAll('#timeline .tl-check').forEach(function(c) {{ c.checked = selCb.checked; }});
                    updateExportBar();
                }};
                var selLabel = document.createElement('label');
                selLabel.className = 'tl-check-label';
                selLabel.setAttribute('for', 'select-all-cb');
                selCbWrap.appendChild(selCb);
                selCbWrap.appendChild(selLabel);
                var selText = document.createElement('span');
                selText.className = 'select-all-label';
                selText.textContent = 'Tout';
                selRow.appendChild(selText);
                selRow.appendChild(selCbWrap);
                selContainer.appendChild(selRow);
            }}
        }}

        function updateExportBar() {{
            var checked = document.querySelectorAll('#timeline .tl-check:checked');
            var bar = document.getElementById('export-bar');
            if (checked.length > 0) {{
                bar.classList.add('visible');
                document.getElementById('export-btn').textContent =
                    'Exporter ' + checked.length + ' planning' + (checked.length > 1 ? 's' : '');
            }} else {{
                bar.classList.remove('visible');
            }}
        }}

        function pad2(n) {{ return n.toString().padStart(2, '0'); }}

        function toICSDate(dt) {{
            return dt.getFullYear().toString() +
                pad2(dt.getMonth() + 1) + pad2(dt.getDate()) + 'T' +
                pad2(dt.getHours()) + pad2(dt.getMinutes()) + '00';
        }}

        function icsEscape(str) {{
            return str.replace(/\\\\/g, '\\\\\\\\').replace(/\\n/g, '\\\\n').replace(/,/g, '\\\\,').replace(/;/g, '\\\\;');
        }}

        function generateICSForNames(names) {{
            // Build notes description from NOTES_DATA
            var noteDesc = '';
            if (NOTES_DATA.comment) {{
                noteDesc += '\\n---\\n' + NOTES_DATA.comment;
            }}
            (NOTES_DATA.updates || []).forEach(function(u) {{
                if (u.text) {{
                    var prefix = u.date ? ('MAJ ' + formatDateFR(u.date) + ': ') : 'MAJ: ';
                    noteDesc += '\\n' + prefix + u.text;
                }}
            }});

            var lines = [
                'BEGIN:VCALENDAR', 'VERSION:2.0',
                'PRODID:-//Planning Urban 7D//FR',
                'CALSCALE:GREGORIAN', 'METHOD:PUBLISH',
                'X-WR-CALNAME:Planning Urban 7D',
                'X-WR-TIMEZONE:Europe/Paris'
            ];
            names.forEach(function(name) {{
                var emp = DATA[name];
                if (!emp) return;
                emp.events.forEach(function(ev, i) {{
                    var s = new Date(ev.start);
                    var e = new Date(ev.end);
                    var desc = ev.label + noteDesc;
                    lines.push('BEGIN:VEVENT');
                    lines.push('UID:export-' + emp.slug + '-' + i + '@urban7d');
                    lines.push('DTSTART;TZID=Europe/Paris:' + toICSDate(s));
                    lines.push('DTEND;TZID=Europe/Paris:' + toICSDate(e));
                    lines.push('SUMMARY:' + name + ' - ' + ev.label);
                    lines.push('DESCRIPTION:' + icsEscape(desc));
                    lines.push('END:VEVENT');
                }});
            }});
            lines.push('END:VCALENDAR');
            return lines.join('\\r\\n');
        }}

        document.getElementById('export-btn').onclick = function() {{
            var checked = document.querySelectorAll('#timeline .tl-check:checked');
            var names = [];
            checked.forEach(function(c) {{ names.push(c.getAttribute('data-name')); }});
            if (names.length === 0) return;
            var ics = generateICSForNames(names);
            var blob = new Blob([ics], {{ type: 'text/calendar;charset=utf-8' }});
            var url = URL.createObjectURL(blob);
            var a = document.createElement('a');
            a.href = url;
            a.download = 'planning-export.ics';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }};

        // ── View toggle ──
        document.querySelectorAll('.view-btn').forEach(function(btn) {{
            btn.onclick = function() {{
                currentView = btn.getAttribute('data-view');
                document.querySelectorAll('.view-btn').forEach(function(b) {{
                    b.classList.toggle('active', b === btn);
                }});
                document.getElementById('view-day').style.display = currentView === 'day' ? '' : 'none';
                document.getElementById('view-staff').style.display = currentView === 'staff' ? '' : 'none';
            }};
        }});

        // ── Modal ──
        var modalEl = document.getElementById('modal');
        document.getElementById('modal-close').onclick = closeModal;
        modalEl.onclick = function(e) {{ if (e.target === modalEl) closeModal(); }};

        function closeModal() {{ modalEl.classList.remove('open'); }}

        function openModal(name) {{
            var emp = DATA[name];
            if (!emp) return;

            document.getElementById('modal-name').textContent = name;
            var body = document.getElementById('modal-body');
            body.innerHTML = '';

            // Group events by day
            var byDay = {{}};
            emp.events.forEach(function(ev) {{ if (!byDay[ev.day]) byDay[ev.day] = []; byDay[ev.day].push(ev); }});

            var hasDays = false;
            for (var d = 0; d < 7; d++) {{
                if (!byDay[d] || byDay[d].length === 0) continue;
                hasDays = true;
                var dayDiv = document.createElement('div');
                dayDiv.className = 'modal-day';

                var title = document.createElement('div');
                title.className = 'modal-day-title';
                title.textContent = DAYS_FULL[d];
                dayDiv.appendChild(title);

                byDay[d].forEach(function(ev) {{
                    var c = getColor(ev.code);
                    var s = new Date(ev.start);
                    var e = new Date(ev.end);
                    var evDiv = document.createElement('div');
                    evDiv.className = 'modal-event';
                    evDiv.style.cssText = 'background:' + c.bg + ';border-color:' + c.border +
                        ';box-shadow:0 0 8px ' + c.border + '30;';

                    var timeSpan = document.createElement('span');
                    timeSpan.className = 'ev-time';
                    timeSpan.style.color = c.text;
                    timeSpan.textContent = s.getHours().toString().padStart(2,'0') + ':' +
                        s.getMinutes().toString().padStart(2,'0') + ' \u2192 ' +
                        e.getHours().toString().padStart(2,'0') + ':' +
                        e.getMinutes().toString().padStart(2,'0');

                    var labelSpan = document.createElement('span');
                    labelSpan.className = 'ev-label';
                    labelSpan.style.color = c.text;
                    labelSpan.textContent = ev.label;

                    evDiv.appendChild(timeSpan);
                    evDiv.appendChild(labelSpan);
                    dayDiv.appendChild(evDiv);
                }});
                body.appendChild(dayDiv);
            }}

            if (!hasDays) {{
                body.innerHTML = '<div class="no-events">Repos cette semaine</div>';
            }}

            // Subscribe button (webcal://)
            var subBtn = document.getElementById('modal-subscribe');
            var icsPath = 'ics/' + emp.slug + '.ics';
            if (window.location.protocol === 'https:' || window.location.protocol === 'http:') {{
                var base = window.location.href.replace(/[^/]*$/, '');
                var fullUrl = new URL(icsPath, base);
                subBtn.href = 'webcal://' + fullUrl.host + fullUrl.pathname;
            }} else {{
                subBtn.href = icsPath;
            }}

            modalEl.classList.add('open');
        }}

        // ── Staff list click ──
        document.querySelectorAll('.employee-btn[data-name]').forEach(function(btn) {{
            btn.onclick = function() {{ openModal(btn.getAttribute('data-name')); }};
        }});

        // ── Notes de semaine (injectées depuis notes/SXX.json) ──
        var REPO = 'OhLaPey/planning-urbansoccer';
        var NOTES_PATH = 'notes/S{week_num}.json';
        var TOKEN_KEY = 'planning-admin-token';
        var notesEl = document.getElementById('week-notes');
        var notesWork = JSON.parse(JSON.stringify(NOTES_DATA));
        var notesDirty = false;

        function getToken() {{ return localStorage.getItem(TOKEN_KEY) || ''; }}
        function setToken(t) {{ localStorage.setItem(TOKEN_KEY, t); }}

        function renderNotes() {{
            var data = notesWork;
            notesEl.innerHTML = '';

            // Comment card
            var card = document.createElement('div');
            card.className = 'note-card comment';
            var hdr = document.createElement('div');
            hdr.className = 'note-header';
            hdr.innerHTML = '<span class="note-label comment">Note de semaine</span>';
            var editBtn = document.createElement('button');
            editBtn.className = 'note-btn';
            editBtn.innerHTML = '\u270e';
            editBtn.title = '\u00c9diter';
            hdr.appendChild(editBtn);
            card.appendChild(hdr);
            var txt = document.createElement('div');
            txt.className = 'note-text';
            txt.textContent = data.comment || '';
            card.appendChild(txt);
            notesEl.appendChild(card);

            editBtn.onclick = function() {{
                if (txt.contentEditable === 'true') {{
                    txt.contentEditable = 'false';
                    data.comment = txt.innerText;
                    editBtn.innerHTML = '\u270e';
                    notesDirty = true;
                    renderNotes();
                }} else {{
                    txt.contentEditable = 'true';
                    txt.focus();
                    editBtn.innerHTML = '\u2714';
                }}
            }};

            // Update cards
            data.updates.forEach(function(u, idx) {{
                var ucard = document.createElement('div');
                ucard.className = 'note-card update';
                var uhdr = document.createElement('div');
                uhdr.className = 'note-header';
                var dateLabel = u.date ? (' \u2014 ' + formatDateFR(u.date)) : '';
                uhdr.innerHTML = '<span class="note-label update">Mise \u00e0 jour' + dateLabel + '</span>';
                var uactions = document.createElement('div');
                uactions.className = 'note-actions';
                var uedit = document.createElement('button');
                uedit.className = 'note-btn';
                uedit.innerHTML = '\u270e';
                uedit.title = '\u00c9diter';
                var udel = document.createElement('button');
                udel.className = 'note-btn del';
                udel.innerHTML = '\u2716';
                udel.title = 'Supprimer';
                uactions.appendChild(uedit);
                uactions.appendChild(udel);
                uhdr.appendChild(uactions);
                ucard.appendChild(uhdr);
                var utxt = document.createElement('div');
                utxt.className = 'note-text';
                utxt.textContent = u.text || '';
                ucard.appendChild(utxt);
                notesEl.appendChild(ucard);

                uedit.onclick = function() {{
                    if (utxt.contentEditable === 'true') {{
                        utxt.contentEditable = 'false';
                        data.updates[idx].text = utxt.innerText;
                        uedit.innerHTML = '\u270e';
                        notesDirty = true;
                        renderNotes();
                    }} else {{
                        utxt.contentEditable = 'true';
                        utxt.focus();
                        uedit.innerHTML = '\u2714';
                    }}
                }};
                udel.onclick = function() {{
                    data.updates.splice(idx, 1);
                    notesDirty = true;
                    renderNotes();
                }};
            }});

            // Add update button
            var addBtn = document.createElement('button');
            addBtn.className = 'add-note-btn';
            addBtn.textContent = '+ Ajouter une mise \u00e0 jour';
            addBtn.onclick = function() {{
                var today = new Date();
                var ds = today.getFullYear() + '-' +
                    (today.getMonth()+1).toString().padStart(2,'0') + '-' +
                    today.getDate().toString().padStart(2,'0') + 'T' +
                    today.getHours().toString().padStart(2,'0') + ':' +
                    today.getMinutes().toString().padStart(2,'0');
                data.updates.push({{ date: ds, text: '' }});
                notesDirty = true;
                renderNotes();
                var cards = notesEl.querySelectorAll('.note-card.update .note-text');
                if (cards.length > 0) {{
                    var last = cards[cards.length - 1];
                    last.contentEditable = 'true';
                    last.focus();
                    var editBtns = notesEl.querySelectorAll('.note-card.update .note-btn:not(.del)');
                    if (editBtns.length > 0) editBtns[editBtns.length - 1].innerHTML = '\u2714';
                }}
            }};
            notesEl.appendChild(addBtn);

            // Publish button (visible when changes exist or token not set)
            var token = getToken();
            if (!token) {{
                // Admin setup: enter GitHub token
                var setup = document.createElement('div');
                setup.className = 'admin-setup';
                var inp = document.createElement('input');
                inp.className = 'admin-input';
                inp.type = 'password';
                inp.placeholder = 'Token GitHub (admin)';
                var saveBtn = document.createElement('button');
                saveBtn.className = 'admin-save-btn';
                saveBtn.textContent = 'Enregistrer';
                saveBtn.onclick = function() {{
                    if (inp.value.trim()) {{
                        setToken(inp.value.trim());
                        renderNotes();
                    }}
                }};
                setup.appendChild(inp);
                setup.appendChild(saveBtn);
                notesEl.appendChild(setup);
                var hint = document.createElement('div');
                hint.className = 'admin-hint';
                hint.textContent = 'Token requis pour publier les notes (GitHub > Settings > Tokens > contents: write)';
                notesEl.appendChild(hint);
            }} else if (notesDirty) {{
                var pubBtn = document.createElement('button');
                pubBtn.className = 'publish-btn';
                pubBtn.textContent = 'Publier les notes';
                pubBtn.onclick = function() {{
                    pubBtn.disabled = true;
                    pubBtn.textContent = 'Publication en cours...';
                    pushNotesToGitHub(data, pubBtn);
                }};
                notesEl.appendChild(pubBtn);
            }}
        }}

        function pushNotesToGitHub(data, btn) {{
            var token = getToken();
            var content = btoa(unescape(encodeURIComponent(JSON.stringify(data, null, 2) + '\\n')));
            var apiUrl = 'https://api.github.com/repos/' + REPO + '/contents/' + NOTES_PATH;

            // First get the current file SHA (required for update)
            fetch(apiUrl, {{
                headers: {{ 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github.v3+json' }}
            }})
            .then(function(r) {{ return r.ok ? r.json() : {{ sha: null }}; }})
            .then(function(file) {{
                var body = {{
                    message: 'MAJ notes S{week_num} depuis la page',
                    content: content,
                    branch: 'main'
                }};
                if (file.sha) body.sha = file.sha;

                return fetch(apiUrl, {{
                    method: 'PUT',
                    headers: {{
                        'Authorization': 'Bearer ' + token,
                        'Accept': 'application/vnd.github.v3+json',
                        'Content-Type': 'application/json'
                    }},
                    body: JSON.stringify(body)
                }});
            }})
            .then(function(r) {{
                if (r.ok) {{
                    notesDirty = false;
                    showRefreshCountdown(btn);
                }} else {{
                    return r.json().then(function(err) {{
                        btn.disabled = false;
                        btn.textContent = 'Erreur : ' + (err.message || 'v\u00e9rifier le token');
                        btn.classList.remove('success');
                    }});
                }}
            }})
            .catch(function(e) {{
                btn.disabled = false;
                btn.textContent = 'Erreur r\u00e9seau, r\u00e9essayer';
            }});
        }}

        function showRefreshCountdown(btn) {{
            var seconds = 90;
            btn.classList.add('success');
            btn.disabled = true;

            function tick() {{
                if (seconds > 0) {{
                    btn.textContent = 'Publi\u00e9 \u2714 En ligne dans ~' + seconds + 's \u2014 Rafra\u00eechir';
                    seconds--;
                    setTimeout(tick, 1000);
                }} else {{
                    btn.textContent = "C'est en ligne ! Rafra\u00eechir la page";
                }}
                btn.disabled = false;
                btn.onclick = function() {{ location.reload(); }};
            }}
            tick();
        }}

        renderNotes();

        // Initial render
        renderTimeline();

        // ── Admin edit mode ──
        var editMode = false;
        var adminToolbarEl = null;

        function initAdminToolbar() {{
            if (!getToken()) return;
            if (adminToolbarEl) return;
            adminToolbarEl = document.createElement('div');
            adminToolbarEl.className = 'admin-toolbar';
            adminToolbarEl.innerHTML = '<span class="label">Admin</span>';
            var toggleBtn = document.createElement('button');
            toggleBtn.className = 'edit-toggle';
            toggleBtn.textContent = 'Mode \u00e9dition';
            toggleBtn.onclick = function() {{
                editMode = !editMode;
                toggleBtn.classList.toggle('active', editMode);
                toggleBtn.textContent = editMode ? 'Quitter \u00e9dition' : 'Mode \u00e9dition';
                renderTimeline();
            }};
            adminToolbarEl.appendChild(toggleBtn);
            var statusEl = document.createElement('span');
            statusEl.className = 'edit-status';
            statusEl.id = 'edit-status';
            adminToolbarEl.appendChild(statusEl);
            var viewDay = document.getElementById('view-day');
            viewDay.insertBefore(adminToolbarEl, viewDay.firstChild);
        }}

        // Override renderTimeline to add editable class when editMode
        var _origRenderTimeline = renderTimeline;
        renderTimeline = function() {{
            _origRenderTimeline();
            if (!editMode) return;
            document.querySelectorAll('#timeline .tl-bar').forEach(function(bar) {{
                bar.classList.add('editable');
            }});
            // Add click handlers for editing
            var rows = document.querySelectorAll('#timeline .timeline-row');
            rows.forEach(function(row) {{
                var nameEl = row.querySelector('.tl-name');
                if (!nameEl || !nameEl.title) return;
                var empName = nameEl.title;
                row.querySelectorAll('.tl-bar').forEach(function(bar, idx) {{
                    bar.onclick = function(e) {{
                        if (!editMode) return;
                        e.stopPropagation();
                        var emp = DATA[empName];
                        if (!emp) return;
                        var dayEvts = emp.events.filter(function(ev) {{ return ev.day === currentDay; }});
                        if (!dayEvts[idx]) return;
                        openEditPopup(empName, dayEvts[idx], idx);
                    }};
                }});
            }});
        }};

        function openEditPopup(empName, ev, evIdx) {{
            // Remove existing popup
            var old = document.getElementById('edit-overlay');
            if (old) old.remove();
            old = document.getElementById('edit-popup');
            if (old) old.remove();

            var s = new Date(ev.start);
            var e = new Date(ev.end);
            var sh = s.getHours().toString().padStart(2, '0') + ':' + s.getMinutes().toString().padStart(2, '0');
            var eh = e.getHours().toString().padStart(2, '0') + ':' + e.getMinutes().toString().padStart(2, '0');

            var overlay = document.createElement('div');
            overlay.className = 'edit-overlay';
            overlay.id = 'edit-overlay';
            overlay.onclick = function() {{ closeEditPopup(); }};
            document.body.appendChild(overlay);

            var popup = document.createElement('div');
            popup.className = 'edit-popup';
            popup.id = 'edit-popup';
            popup.innerHTML =
                '<h3>' + empName + ' \u2014 ' + ev.label + '</h3>' +
                '<div class="field"><label>D\u00e9but</label><input type="time" id="edit-start" value="' + sh + '"></div>' +
                '<div class="field"><label>Fin</label><input type="time" id="edit-end" value="' + eh + '"></div>' +
                '<div class="actions">' +
                '<button class="btn-cancel" id="edit-cancel">Annuler</button>' +
                '<button class="btn-save" id="edit-save">Enregistrer</button>' +
                '</div>';
            document.body.appendChild(popup);

            document.getElementById('edit-cancel').onclick = closeEditPopup;
            document.getElementById('edit-save').onclick = function() {{
                var newStart = document.getElementById('edit-start').value;
                var newEnd = document.getElementById('edit-end').value;
                if (!newStart || !newEnd) return;
                applyTimeEdit(empName, ev, newStart, newEnd);
                closeEditPopup();
            }};
        }}

        function closeEditPopup() {{
            var el = document.getElementById('edit-overlay');
            if (el) el.remove();
            el = document.getElementById('edit-popup');
            if (el) el.remove();
        }}

        function applyTimeEdit(empName, ev, newStart, newEnd) {{
            // Update the DATA object in memory
            var dateStr = ev.start.substring(0, 11); // "2026-03-02T"
            ev.start = dateStr + newStart;
            ev.end = dateStr + newEnd;
            renderTimeline();

            // Show saving status
            var statusEl = document.getElementById('edit-status');
            if (statusEl) statusEl.textContent = 'Sauvegarde...';

            // Push updated data to GitHub
            pushDataToGitHub(function(ok) {{
                if (statusEl) {{
                    statusEl.textContent = ok ? 'Sauvegard\u00e9 \u2714' : 'Erreur !';
                    setTimeout(function() {{ statusEl.textContent = ''; }}, 3000);
                }}
            }});
        }}

        function pushDataToGitHub(cb) {{
            var token = getToken();
            if (!token) {{ cb(false); return; }}

            // Build the updated data JSON for this week
            var weekData = {{}};
            Object.keys(DATA).forEach(function(name) {{
                if (name === '_codeNames') return;
                weekData[name] = DATA[name];
            }});
            var content = btoa(unescape(encodeURIComponent(JSON.stringify(weekData, null, 2) + '\\n')));
            var dataPath = 'data/S{week_num}-events.json';
            var apiUrl = 'https://api.github.com/repos/' + REPO + '/contents/' + dataPath;

            // Get current SHA
            fetch(apiUrl, {{
                headers: {{ 'Authorization': 'Bearer ' + token, 'Accept': 'application/vnd.github.v3+json' }}
            }})
            .then(function(r) {{ return r.ok ? r.json() : {{ sha: null }}; }})
            .then(function(file) {{
                var body = {{
                    message: 'MAJ cr\u00e9neaux S{week_num} depuis la page',
                    content: content,
                    branch: 'main'
                }};
                if (file.sha) body.sha = file.sha;

                return fetch(apiUrl, {{
                    method: 'PUT',
                    headers: {{
                        'Authorization': 'Bearer ' + token,
                        'Accept': 'application/vnd.github.v3+json',
                        'Content-Type': 'application/json'
                    }},
                    body: JSON.stringify(body)
                }});
            }})
            .then(function(r) {{ cb(r.ok); }})
            .catch(function() {{ cb(false); }});
        }}

        // Init admin toolbar if token exists
        initAdminToolbar();
        // Re-init after token is saved in notes section
        var _origSetToken = setToken;
        setToken = function(t) {{
            _origSetToken(t);
            initAdminToolbar();
        }};

    }})();
    </script>
</body>
</html>"""


# ── Main ───────────────────────────────────────────────────────────────────


def main():
    excel_files = discover_excel_files()
    if not excel_files:
        print("Aucun fichier 'Plannings YYYY SXX.xlsx' trouv\u00e9.")
        return

    print(f"Fichiers Excel trouv\u00e9s : {len(excel_files)}")
    for ef in excel_files:
        print(f"  - {ef['filename']} (S{ef['week']}, {ef['year']})")

    # ── Collecter tous les événements par employé, toutes semaines ──
    all_employee_events = {}   # {name: [events]}
    week_data = {}             # {week_num: {employees, year}}
    all_weeks = set()

    for ef in excel_files:
        year, week_num = ef["year"], ef["week"]
        dates = week_dates(year, week_num)

        wb = openpyxl.load_workbook(ef["filename"])
        ws = wb["Planning"] if "Planning" in wb.sheetnames else wb.active

        employees = parse_employees(ws, dates, week_num)
        all_weeks.add(week_num)
        week_data[week_num] = {"employees": employees, "year": year}

        active_count = 0
        print(f"\nSemaine {week_num} ({year}) :")
        for name, evts in employees.items():
            if name not in all_employee_events:
                all_employee_events[name] = []
            all_employee_events[name].extend(evts)
            if evts:
                active_count += 1
                print(f"  {name} ({len(evts)} \u00e9v\u00e9nements)")
                for e in evts:
                    end_str = e["end"].strftime("%H:%M")
                    if e["end"].date() > e["start"].date():
                        end_str += " (+1j)"
                    print(f"    {e['start'].strftime('%a %d/%m %H:%M')} - "
                          f"{end_str} : {e['label']}")
        print(f"  \u2192 {active_count} employ\u00e9s actifs")

    # ── Charger les notes par semaine ──
    all_week_notes = {}
    for wn in all_weeks:
        all_week_notes[wn] = load_week_notes(wn)

    # ── Générer les fichiers ICS (cumulatifs, toutes semaines) ──
    os.makedirs("ics", exist_ok=True)
    ics_count = 0
    for name, events in all_employee_events.items():
        if events:
            events.sort(key=lambda e: e["start"])
            ics_content = generate_ics(name, events, week_notes=all_week_notes)
            filename = f"ics/{slug(name)}.ics"
            with open(filename, "w", encoding="utf-8") as f:
                f.write(ics_content)
            ics_count += 1
    print(f"\n{ics_count} fichiers ICS g\u00e9n\u00e9r\u00e9s dans ics/")

    # ── Générer HTML + JSON par semaine ──
    os.makedirs("data", exist_ok=True)
    for week_num in sorted(all_weeks):
        wd = week_data[week_num]
        year = wd["year"]
        employees = wd["employees"]

        # JSON
        active_names = sorted([n for n, e in employees.items() if e])
        monday = datetime.fromisocalendar(year, week_num, 1)
        sunday = monday + timedelta(days=6)
        json_data = {
            "semaine": week_num,
            "annee": year,
            "date_debut": f"{monday.day} {FRENCH_MONTHS[monday.month]}",
            "date_fin": f"{sunday.day} {FRENCH_MONTHS[sunday.month]}",
            "employesActifs": active_names,
        }
        json_path = f"data/S{week_num}.json"
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(json_data, f, ensure_ascii=False, indent=2)
        print(f"\u00c9crit : {json_path}")

        # HTML
        html_content = generate_html(employees, week_num, year, all_weeks)
        html_path = f"S{week_num}.html"
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        print(f"\u00c9crit : {html_path}")

    # ── Mettre à jour index.html → dernière semaine ──
    latest_week = max(all_weeks)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(
            '<!DOCTYPE html>\n'
            '<html lang="fr">\n'
            '<head>\n'
            '    <meta charset="UTF-8">\n'
            f'    <meta http-equiv="refresh" content="0;url=S{latest_week}.html">\n'
            '    <title>Planning Urban 7D</title>\n'
            '</head>\n'
            '<body>\n'
            f'    <p>Redirection vers <a href="S{latest_week}.html">S{latest_week}</a>...</p>\n'
            '</body>\n'
            '</html>'
        )
    print(f"\u00c9crit : index.html \u2192 S{latest_week}.html")

    print("\nTermin\u00e9 !")
    print("\n\u2500\u2500 Abonnement calendrier \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500")
    print("1. H\u00e9bergez ces fichiers (GitHub Pages, Netlify, etc.)")
    print("2. Les employ\u00e9s ouvrent la page et cliquent sur leur nom")
    print("3. Le calendrier se met \u00e0 jour automatiquement")
    print(f"\nPour ajouter une nouvelle semaine :")
    print(f"  1. Ajoutez le fichier Excel « Plannings {year} SXX.xlsx »")
    print(f"  2. Relancez : python generate.py")
    print(f"  3. Publiez les fichiers mis \u00e0 jour")


if __name__ == "__main__":
    main()

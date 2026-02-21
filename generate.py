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

# ── Mapping codes → noms lisibles ──────────────────────────────────────────

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

COLS = ["B", "C", "D", "E", "F", "G", "H"]

FRENCH_MONTHS = {
    1: "Janvier", 2: "Février", 3: "Mars", 4: "Avril",
    5: "Mai", 6: "Juin", 7: "Juillet", 8: "Août",
    9: "Septembre", 10: "Octobre", 11: "Novembre", 12: "Décembre",
}

# ── Utilitaires ────────────────────────────────────────────────────────────


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
    pattern = re.compile(r"Plannings\s+(\d{4})\s+S(\d+)\.xlsx", re.IGNORECASE)
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
                employees[current_name] = parse_shifts(ws, current_rows, dates, week_num)
            current_name = name_cell.strip()
            current_rows = [row]
        elif current_name:
            current_rows.append(row)

    if current_name:
        employees[current_name] = parse_shifts(ws, current_rows, dates, week_num)

    return employees


def parse_shifts(ws, rows, dates, week_num):
    """Parse les créneaux d'un employé à partir de ses lignes."""
    events = []

    i = 0
    while i < len(rows):
        row = rows[i]
        has_codes = False
        codes = {}
        for col in COLS:
            val = get_cell(ws, row, col)
            if val and isinstance(val, str):
                val = val.strip()
                if val and not re.match(r"^\d{2}:\d{2}/\d{2}:\d{2}", val):
                    has_codes = True
                    codes[col] = val

        if has_codes:
            times = {}
            if i + 1 < len(rows):
                time_row = rows[i + 1]
                for col in COLS:
                    val = get_cell(ws, time_row, col)
                    if val and isinstance(val, str) and re.match(r"^\d{2}:\d{2}/\d{2}:\d{2}", val.strip()):
                        times[col] = val.strip()

            for col, code in codes.items():
                if col in times and col in dates:
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

            i += 2
        else:
            i += 1

    events.sort(key=lambda e: e["start"])
    return events


# ── Génération ICS (abonnement calendrier) ─────────────────────────────────


def generate_ics(name, events):
    """Génère le contenu ICS pour un employé (toutes semaines confondues).

    Chaque fichier ICS contient TOUS les événements de l'employé, ce qui
    permet à l'abonnement calendrier de rester à jour automatiquement.
    """
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
        for i, evt in enumerate(by_week[week_num], 1):
            dt_start = evt["start"].strftime("%Y%m%dT%H%M%S")
            dt_end = evt["end"].strftime("%Y%m%dT%H%M%S")
            lines.extend([
                "BEGIN:VEVENT",
                f"UID:{s}-s{week_num}-{i}@urban7d",
                f"DTSTAMP:{dt_start}",
                f"DTSTART;TZID=Europe/Paris:{dt_start}",
                f"DTEND;TZID=Europe/Paris:{dt_end}",
                f"SUMMARY:{evt['label']}",
                f"DESCRIPTION:{evt['label']}",
                "END:VEVENT",
            ])

    lines.append("END:VCALENDAR")
    return "\r\n".join(lines)


# ── Génération HTML ────────────────────────────────────────────────────────


def generate_html(week_employees, week_num, year, all_weeks):
    """Génère la page HTML pour une semaine (avec liens d'abonnement)."""
    active_names = [name for name, evts in week_employees.items() if len(evts) > 0]
    date_range = format_date_range(year, week_num)

    week_tabs = ""
    for w in sorted(all_weeks):
        if w == week_num:
            week_tabs += f'            <a href="#" class="week-tab active">S{w}</a>\n'
        else:
            week_tabs += f'            <a href="S{w}.html" class="week-tab">S{w}</a>\n'

    employee_lines = ""
    for name in week_employees:
        s = slug(name)
        if name in active_names:
            employee_lines += (
                f'            <a href="ics/{s}.ics" class="employee" '
                f'data-ics="ics/{s}.ics">{name}</a>\n'
            )
        else:
            employee_lines += (
                f'            <div class="employee repos">{name} '
                f'<span class="badge">Repos</span></div>\n'
            )

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Planning Urban 7D - S{week_num}</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{ font-family: 'Inter', sans-serif; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); min-height: 100vh; padding: 20px; }}
        .container {{ max-width: 500px; margin: 0 auto; }}
        .header {{ text-align: center; margin-bottom: 20px; padding: 20px; }}
        .logo {{ font-size: 48px; margin-bottom: 10px; }}
        h1 {{ color: #FF6B35; font-size: 28px; font-weight: 700; margin-bottom: 8px; }}
        .subtitle {{ color: #888; font-size: 14px; margin-bottom: 5px; }}
        .dates {{ color: #FF6B35; font-size: 18px; font-weight: 600; background: rgba(255, 107, 53, 0.1); padding: 10px 20px; border-radius: 20px; display: inline-block; margin-top: 10px; }}
        .week-selector {{ display: flex; justify-content: center; gap: 8px; margin-bottom: 25px; flex-wrap: wrap; }}
        .week-tab {{ padding: 10px 18px; background: rgba(255, 255, 255, 0.05); border: 1px solid rgba(255, 255, 255, 0.1); border-radius: 25px; color: #888; text-decoration: none; font-weight: 500; font-size: 14px; transition: all 0.2s ease; }}
        .week-tab:hover {{ background: rgba(255, 107, 53, 0.1); border-color: rgba(255, 107, 53, 0.3); color: #FF6B35; }}
        .week-tab.active {{ background: #FF6B35; border-color: #FF6B35; color: white; }}
        .employees {{ display: flex; flex-direction: column; gap: 8px; }}
        .employee {{ display: flex; align-items: center; justify-content: space-between; padding: 16px 20px; background: rgba(255, 255, 255, 0.05); border-radius: 12px; color: white; text-decoration: none; font-weight: 500; transition: all 0.2s ease; border: 1px solid rgba(255, 255, 255, 0.1); }}
        a.employee:hover {{ background: rgba(255, 107, 53, 0.2); border-color: #FF6B35; transform: translateX(5px); }}
        a.employee::after {{ content: '\U0001F4C5'; font-size: 20px; }}
        .employee.repos {{ color: #555; background: rgba(255, 255, 255, 0.02); border-color: rgba(255, 255, 255, 0.05); }}
        .badge {{ font-size: 11px; padding: 4px 10px; background: rgba(255, 255, 255, 0.1); border-radius: 20px; color: #555; font-weight: 400; }}
        .footer {{ text-align: center; margin-top: 30px; color: #666; font-size: 12px; line-height: 1.8; }}
        .footer a {{ color: #FF6B35; text-decoration: none; }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="logo">\u26bd</div>
            <h1>Planning Urban 7D</h1>
            <p class="subtitle">Semaine {week_num}</p>
            <div class="dates">{date_range}</div>
        </div>
        <div class="week-selector">
{week_tabs.rstrip()}
        </div>
        <div class="employees">
{employee_lines.rstrip()}
        </div>
        <div class="footer">
            <p>Cliquez sur votre nom pour vous abonner au calendrier</p>
            <p style="margin-top: 4px; font-size: 11px; color: #555;">
                L'abonnement se met \u00e0 jour automatiquement \u00e0 chaque nouvelle semaine
            </p>
        </div>
    </div>
    <script>
        // Convertir les liens ICS en webcal:// pour l'abonnement calendrier.
        // webcal:// ouvre directement l'app calendrier (iOS, macOS, Android, Outlook).
        // En local (file://), on garde le lien direct pour le téléchargement.
        if (window.location.protocol === 'https:' || window.location.protocol === 'http:') {{
            document.querySelectorAll('a.employee[data-ics]').forEach(function(a) {{
                var icsPath = a.getAttribute('data-ics');
                var base = window.location.href.replace(/[^/]*$/, '');
                var fullUrl = new URL(icsPath, base);
                a.href = 'webcal://' + fullUrl.host + fullUrl.pathname;
            }});
        }}
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

    # ── Générer les fichiers ICS (cumulatifs, toutes semaines) ──
    os.makedirs("ics", exist_ok=True)
    ics_count = 0
    for name, events in all_employee_events.items():
        if events:
            events.sort(key=lambda e: e["start"])
            ics_content = generate_ics(name, events)
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

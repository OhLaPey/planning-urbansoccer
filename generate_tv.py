#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Génère la page « Programme TV » du centre (tv.html) à partir de
data/tv-programme.json.

Même logique que les plannings :
    data/tv-programme.json  →  generate_tv.py  →  tv.html (statique, GitHub Pages)

La page n'affiche comme DIFFUSABLE que les événements dont la chaîne fait
partie de l'abonnement du centre (Canal+, beIN Sports, chaînes en clair…).
Les événements sur des chaînes non disponibles (Ligue 1+, DAZN) sont
regroupés à part, en grisé, pour que l'équipe sache pourquoi un match
n'est pas proposé à la diffusion.

Usage :
    python generate_tv.py
"""

import json
import os
from datetime import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(HERE, "data", "tv-programme.json")
OUTPUT_PATH = os.path.join(HERE, "tv.html")

FRENCH_DAYS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
FRENCH_MONTHS = ["", "janvier", "février", "mars", "avril", "mai", "juin",
                 "juillet", "août", "septembre", "octobre", "novembre", "décembre"]


def load_data():
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def build_page(data):
    """Construit le HTML. Tout le rendu dynamique (aujourd'hui, EN DIRECT,
    tri, regroupement par jour) est fait côté client à partir de DATA, pour
    que la page reste juste sans régénération pendant la journée."""

    meta = data.get("_meta", {})
    titre = meta.get("titre", "Programme TV")
    sous_titre = meta.get("sous_titre", "")
    updated_at = meta.get("updated_at", "")

    # On ne pousse dans la page que ce dont le client a besoin.
    payload = {
        "meta": meta,
        "abonnement": data.get("abonnement", {"disponibles": [], "non_disponibles": []}),
        "chaines_meta": data.get("chaines_meta", {}),
        "categories": data.get("categories", {}),
        "evenements": data.get("evenements", []),
    }
    data_json = json.dumps(payload, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="apple-mobile-web-app-capable" content="yes">
    <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
    <meta name="apple-mobile-web-app-title" content="TV Centre">
    <meta name="theme-color" content="#FF6600">
    <title>Programme TV — Centre UrbanSoccer</title>
    <link rel="apple-touch-icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>&#128250;</text></svg>">
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: 'Montserrat', sans-serif;
            background: #1A1A1A;
            min-height: 100vh;
            padding: 15px;
            color: #fff;
            position: relative;
        }}
        body::before {{
            content: '';
            position: fixed; inset: 0; z-index: 0; pointer-events: none;
            background: url('bg-team.jpg') center center / cover no-repeat;
            opacity: 0.18;
        }}
        .container {{
            position: relative; z-index: 1; max-width: 1100px; margin: 0 auto;
            background: rgba(26,26,26,0.95); border-radius: 8px;
            padding: 4px 14px 22px; margin-top: 6px; margin-bottom: 6px;
            border-top: 4px solid #FF6600; overflow: hidden;
        }}
        .container::after {{
            content: '\\276F\\276F\\276F\\276F'; position: absolute;
            top: 14px; right: -2px; font-size: 34px; font-weight: 900;
            color: rgba(255,102,0,0.08); letter-spacing: -5px;
            pointer-events: none; z-index: 0;
        }}

        /* ── Header ── */
        .header {{ text-align: center; padding: 22px 10px 10px; }}
        .header .tv-icon {{ font-size: 24px; }}
        h1 {{
            color: #fff; font-size: 30px; font-weight: 900; margin: 4px 0 2px;
            text-transform: uppercase; letter-spacing: 3px;
            text-shadow: 0 0 30px rgba(255,102,0,0.3);
        }}
        .subtitle {{
            color: #FF6600; font-size: 12px; font-weight: 700;
            text-transform: uppercase; letter-spacing: 4px;
        }}
        .clock {{
            margin-top: 12px; font-size: 15px; color: #bbb; font-weight: 600;
            text-transform: capitalize;
        }}
        .clock .time {{ color: #FF6600; font-weight: 800; }}

        /* ── Légende chaînes ── */
        .legend {{
            display: flex; gap: 8px; flex-wrap: wrap; justify-content: center;
            margin: 14px 0 8px; padding: 12px; border-radius: 8px;
            background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06);
        }}
        .legend-title {{
            width: 100%; text-align: center; font-size: 10px; font-weight: 700;
            color: #777; text-transform: uppercase; letter-spacing: 1.5px;
            margin-bottom: 6px;
        }}
        .chan-badge {{
            display: inline-flex; align-items: center; gap: 7px;
            font-size: 11px; font-weight: 800; padding: 4px 11px 4px 5px; border-radius: 5px;
            color: #fff; text-transform: uppercase; letter-spacing: 0.3px;
            border: 1px solid rgba(255,255,255,0.14); white-space: nowrap;
        }}
        .chan-num {{
            background: rgba(255,255,255,0.95); color: #111; font-weight: 900;
            font-size: 12px; min-width: 22px; text-align: center;
            padding: 2px 5px; border-radius: 4px; letter-spacing: 0;
        }}
        .chan-badge.clair .chan-num {{ background: #FF6600; color: #1A1A1A; }}
        .chan-badge.clair::after {{
            content: 'CLAIR'; font-size: 8px; opacity: 0.85;
            background: rgba(255,255,255,0.2); padding: 1px 4px; border-radius: 3px;
        }}

        /* ── Jour ── */
        .day-block {{ margin-top: 20px; }}
        .day-head {{
            display: flex; align-items: baseline; gap: 10px;
            padding: 8px 4px; border-bottom: 2px solid rgba(255,102,0,0.35);
            margin-bottom: 10px;
        }}
        .day-head .day-name {{
            font-size: 18px; font-weight: 900; text-transform: uppercase;
            letter-spacing: 1px; color: #fff;
        }}
        .day-head .day-date {{ font-size: 12px; color: #888; font-weight: 600; }}
        .day-head.today .day-name {{ color: #FF6600; }}
        .day-head .today-tag {{
            margin-left: auto; font-size: 10px; font-weight: 800;
            background: #FF6600; color: #1A1A1A; padding: 3px 10px;
            border-radius: 5px; text-transform: uppercase; letter-spacing: 1px;
        }}

        /* ── Événement ── */
        .event {{
            display: grid;
            grid-template-columns: 66px 6px 1fr auto;
            align-items: center; gap: 12px;
            padding: 12px 14px; margin-bottom: 8px; border-radius: 8px;
            background: rgba(255,255,255,0.04);
            border: 1px solid rgba(255,255,255,0.07);
            transition: background 0.15s;
        }}
        .event:hover {{ background: rgba(255,255,255,0.06); }}
        .event.live {{
            border-color: rgba(255,102,0,0.55);
            background: rgba(255,102,0,0.08);
            box-shadow: 0 0 22px rgba(255,102,0,0.12);
        }}
        .event.done {{ opacity: 0.42; }}
        .ev-time {{ font-size: 20px; font-weight: 800; color: #fff; text-align: center; }}
        .ev-time small {{ display: block; font-size: 9px; color: #888; font-weight: 700; }}
        .ev-bar {{ width: 6px; height: 100%; min-height: 40px; border-radius: 3px; }}
        .ev-main {{ min-width: 0; }}
        .ev-compet {{
            font-size: 10px; font-weight: 800; text-transform: uppercase;
            letter-spacing: 0.8px; margin-bottom: 3px;
        }}
        .ev-affiche {{
            font-size: 17px; font-weight: 700; color: #fff;
            overflow: hidden; text-overflow: ellipsis;
        }}
        .ev-right {{ display: flex; flex-direction: column; align-items: flex-end; gap: 6px; }}
        .live-tag {{
            font-size: 10px; font-weight: 900; color: #1A1A1A; background: #FF6600;
            padding: 3px 9px; border-radius: 5px; text-transform: uppercase;
            letter-spacing: 1px; display: inline-flex; align-items: center; gap: 5px;
        }}
        .live-tag .dot {{
            width: 7px; height: 7px; border-radius: 50%; background: #1A1A1A;
            animation: pulse 1.1s infinite;
        }}
        @keyframes pulse {{ 0%,100% {{ opacity: 1; }} 50% {{ opacity: 0.25; }} }}

        /* ── Section non diffusable ── */
        .unavailable {{
            margin-top: 30px; padding: 16px; border-radius: 8px;
            background: rgba(255,255,255,0.02);
            border: 1px dashed rgba(255,255,255,0.12);
        }}
        .unavailable h2 {{
            font-size: 12px; font-weight: 800; color: #888; text-transform: uppercase;
            letter-spacing: 1px; margin-bottom: 4px;
        }}
        .unavailable .hint {{ font-size: 11px; color: #666; margin-bottom: 12px; }}
        .unavailable .event {{ opacity: 0.6; }}
        .unavailable .ev-affiche {{ color: #bbb; }}

        /* ── Vide ── */
        .empty-state {{ text-align: center; padding: 46px 20px; color: #666; }}
        .empty-state .big {{ font-size: 40px; margin-bottom: 10px; }}

        /* ── Footer ── */
        .foot {{
            margin-top: 26px; padding-top: 14px; text-align: center;
            border-top: 1px solid rgba(255,255,255,0.06);
            font-size: 10px; color: #555; letter-spacing: 0.5px;
        }}

        @media (max-width: 620px) {{
            h1 {{ font-size: 22px; letter-spacing: 1.5px; }}
            .event {{ grid-template-columns: 52px 5px 1fr; gap: 9px; }}
            .ev-right {{ grid-column: 1 / -1; flex-direction: row; align-items: center;
                          justify-content: space-between; margin-top: 4px; }}
            .ev-time {{ font-size: 17px; }}
            .ev-affiche {{ font-size: 15px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="tv-icon">📺</div>
            <h1>Programme TV</h1>
            <div class="subtitle">{sous_titre or titre}</div>
            <div class="clock" id="clock"></div>
        </div>

        <div class="legend" id="legend"></div>

        <div id="schedule"></div>

        <div class="foot">
            Mis à jour le {updated_at} · Diffusable = chaîne incluse dans l'abonnement du centre
        </div>
    </div>

    <script>
    var DATA = {data_json};

    var JS_DAYS = ["Dimanche","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi"];
    var JS_MONTHS = ["janvier","février","mars","avril","mai","juin","juillet",
                     "août","septembre","octobre","novembre","décembre"];
    var EVENT_DURATION_MIN = 130; // durée supposée d'un match/événement

    function pad(n) {{ return (n < 10 ? "0" : "") + n; }}
    function ymd(d) {{ return d.getFullYear() + "-" + pad(d.getMonth()+1) + "-" + pad(d.getDate()); }}

    function eventStart(ev) {{
        var p = ev.date.split("-");
        var t = (ev.heure || "00:00").split(":");
        return new Date(+p[0], +p[1]-1, +p[2], +t[0], +(t[1]||0));
    }}

    function isAvailable(chaine) {{
        var dispo = (DATA.abonnement.disponibles || []);
        return dispo.indexOf(chaine) !== -1;
    }}

    function chanColor(chaine) {{
        var m = DATA.chaines_meta[chaine];
        return m ? m.couleur : "#455a64";
    }}
    function chanClair(chaine) {{
        var m = DATA.chaines_meta[chaine];
        return m ? !!m.clair : false;
    }}
    function chanNum(chaine) {{
        var m = DATA.chaines_meta[chaine];
        return (m && m.numero) ? m.numero : "";
    }}
    function chanBadge(chaine) {{
        var num = chanNum(chaine);
        var numHtml = num ? '<span class="chan-num">' + num + '</span>' : '';
        return '<span class="chan-badge' + (chanClair(chaine) ? ' clair' : '') +
               '" style="background:' + chanColor(chaine) + '">' +
               numHtml + '<span>' + chaine + '</span></span>';
    }}
    function catInfo(cat) {{
        return DATA.categories[cat] || {{ label: cat || "Sport", couleur: "#78909C" }};
    }}

    function renderClock() {{
        var now = new Date();
        var el = document.getElementById("clock");
        el.innerHTML = JS_DAYS[now.getDay()] + " " + now.getDate() + " " +
            JS_MONTHS[now.getMonth()] + " · <span class=\\"time\\">" +
            pad(now.getHours()) + ":" + pad(now.getMinutes()) + "</span>";
    }}

    function renderLegend() {{
        var wrap = document.getElementById("legend");
        var html = '<div class="legend-title">Chaînes disponibles au centre · numéro à gauche</div>';
        (DATA.abonnement.disponibles || []).forEach(function(ch) {{
            html += chanBadge(ch);
        }});
        wrap.innerHTML = html;
    }}

    function eventCard(ev, now) {{
        var start = eventStart(ev);
        var end = new Date(start.getTime() + EVENT_DURATION_MIN*60000);
        var state = "soon";
        if (now >= start && now < end) state = "live";
        else if (now >= end) state = "done";

        var cat = catInfo(ev.categorie);
        var dispo = isAvailable(ev.chaine);

        var right = chanBadge(ev.chaine);
        if (state === "live" && dispo) {{
            right = '<span class="live-tag"><span class="dot"></span>En direct</span>' + right;
        }}

        return '<div class="event ' + (state==="live"&&dispo?"live":"") + ' ' +
                (state==="done"?"done":"") + '">' +
            '<div class="ev-time">' + (ev.heure||"") + '</div>' +
            '<div class="ev-bar" style="background:' + cat.couleur + '"></div>' +
            '<div class="ev-main">' +
                '<div class="ev-compet" style="color:' + cat.couleur + '">' +
                    (ev.competition || cat.label) + '</div>' +
                '<div class="ev-affiche">' + (ev.affiche || "") + '</div>' +
            '</div>' +
            '<div class="ev-right">' + right + '</div>' +
        '</div>';
    }}

    function render() {{
        renderClock();
        var now = new Date();
        var todayStr = ymd(now);

        var evts = (DATA.evenements || []).slice().sort(function(a,b) {{
            return eventStart(a) - eventStart(b);
        }});

        // On garde aujourd'hui + futur, on masque le passé (hors événement du jour).
        var dispo = [], indispo = [];
        evts.forEach(function(ev) {{
            if (ev.date < todayStr) return; // jours passés masqués
            if (isAvailable(ev.chaine)) dispo.push(ev); else indispo.push(ev);
        }});

        var schedule = document.getElementById("schedule");

        if (!dispo.length && !indispo.length) {{
            schedule.innerHTML = '<div class="empty-state"><div class="big">📺</div>' +
                '<p>Aucun événement programmé pour le moment.</p></div>';
            return;
        }}

        // Regroupement par jour des événements diffusables
        var html = "";
        var groups = {{}}, order = [];
        dispo.forEach(function(ev) {{
            if (!groups[ev.date]) {{ groups[ev.date] = []; order.push(ev.date); }}
            groups[ev.date].push(ev);
        }});

        order.forEach(function(date) {{
            var p = date.split("-");
            var d = new Date(+p[0], +p[1]-1, +p[2]);
            var isToday = (date === todayStr);
            html += '<div class="day-block"><div class="day-head' +
                (isToday ? " today" : "") + '">' +
                '<span class="day-name">' + JS_DAYS[d.getDay()] + '</span>' +
                '<span class="day-date">' + d.getDate() + ' ' + JS_MONTHS[d.getMonth()] + '</span>' +
                (isToday ? '<span class="today-tag">Aujourd\\'hui</span>' : '') +
                '</div>';
            groups[date].forEach(function(ev) {{ html += eventCard(ev, now); }});
            html += '</div>';
        }});

        // Section « non disponible au centre »
        if (indispo.length) {{
            html += '<div class="unavailable"><h2>⛔ Non disponible au centre</h2>' +
                '<div class="hint">Ces événements passent sur des chaînes hors abonnement (' +
                (DATA.abonnement.non_disponibles || []).join(", ") +
                ') — non diffusables sur les écrans.</div>';
            indispo.forEach(function(ev) {{ html += eventCard(ev, now); }});
            html += '</div>';
        }}

        schedule.innerHTML = html;
    }}

    renderLegend();
    render();
    setInterval(render, 60000);            // rafraîchit l'état EN DIRECT chaque minute
    setInterval(function() {{               // recharge la page 1×/h (nouveau programme éventuel)
        location.reload();
    }}, 3600000);
    </script>
</body>
</html>"""


def main():
    if not os.path.exists(DATA_PATH):
        print(f"Fichier introuvable : {DATA_PATH}")
        return
    data = load_data()
    html = build_page(data)
    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    n = len(data.get("evenements", []))
    dispo = set(data.get("abonnement", {}).get("disponibles", []))
    diffusables = sum(1 for e in data.get("evenements", []) if e.get("chaine") in dispo)
    print(f"Écrit : {OUTPUT_PATH}")
    print(f"  {n} événements ({diffusables} diffusables, {n - diffusables} non diffusables)")


if __name__ == "__main__":
    main()

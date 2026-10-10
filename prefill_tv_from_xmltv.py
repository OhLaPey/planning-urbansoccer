#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Pré-remplissage du programme TV à partir d'un flux XMLTV.

PROTOTYPE — génère un BROUILLON à relire, jamais le fichier final.

Chaîne de traitement :
    flux XMLTV (URL/fichier)  ─┐
    data/tv-programme.json     ├─►  prefill_tv_from_xmltv.py  ─►  data/tv-programme.draft.json
    data/xmltv-mapping.json    ─┘                                   (à relire, élaguer, renommer)

Le script :
  1. télécharge (ou lit) un guide XMLTV (gzip ou xml) ;
  2. ne garde que les chaînes du centre (via le mapping + abonnement.disponibles) ;
  3. garde les programmes des chaînes du centre : TOUT sur les chaînes
     sport (Canal+, beIN, L'Équipe…), uniquement le SPORT sur les chaînes
     généralistes TNT (TF1, France 2/3, M6, W9, TMC) ;
  4. devine la catégorie/genre pour le code couleur ;
  5. écrit le programme au schéma de data/tv-programme.json.

C'est un ASSISTANT de saisie : l'équipe relit le brouillon, retire ce qui n'a
pas d'intérêt, corrige les libellés, puis renomme le brouillon en
data/tv-programme.json et lance generate_tv.py.

⚠️ Les flux XMLTV sont non-officiels (scraping), fragiles et en zone grise côté
   conditions d'utilisation. À utiliser en connaissance de cause.

Usage :
    python prefill_tv_from_xmltv.py --source https://exemple/guide.xml.gz
    python prefill_tv_from_xmltv.py --source guide_local.xml --days 7
"""

import argparse
import gzip
import io
import json
import os
import re
import sys
import urllib.request
from datetime import datetime, timedelta

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
PROG_PATH = os.path.join(DATA_DIR, "tv-programme.json")
MAPPING_PATH = os.path.join(DATA_DIR, "xmltv-mapping.json")
DEFAULT_OUT = os.path.join(DATA_DIR, "tv-programme.draft.json")

# ── Catégories (code → libellé + couleur) écrites dans le fichier de sortie ───
CATEGORIES = {
    "ldc":    {"label": "Ligue des Champions", "couleur": "#6495ED"},
    "coupe":  {"label": "Coupes",              "couleur": "#FFA500"},
    "edf":    {"label": "Équipe de France",    "couleur": "#66BB6A"},
    "foot":   {"label": "Football",            "couleur": "#FF8C42"},
    "padel":  {"label": "Padel",               "couleur": "#BA55D3"},
    "tennis": {"label": "Tennis",              "couleur": "#26C6DA"},
    "rugby":  {"label": "Rugby",               "couleur": "#8D6E63"},
    "f1":     {"label": "Formule 1",           "couleur": "#E10600"},
    "sport":  {"label": "Sport",               "couleur": "#78909C"},
    "film":   {"label": "Film",                "couleur": "#5C6BC0"},
    "serie":  {"label": "Série",               "couleur": "#7E57C2"},
    "info":   {"label": "Info",                "couleur": "#42A5F5"},
    "mag":    {"label": "Divertissement",      "couleur": "#EC407A"},
    "doc":    {"label": "Documentaire",        "couleur": "#9CCC65"},
    "jeunesse": {"label": "Jeunesse",          "couleur": "#FFCA28"},
    "culture": {"label": "Culture",            "couleur": "#26A69A"},
    "autre":  {"label": "Programme",           "couleur": "#90A4AE"},
}

# ── Sous-catégories sport (affinage par mots-clés du titre) ───────────────────
SPORT_RULES = [
    ("ldc",    [r"ligue des champions", r"champions league", r"\bc1\b", r"\bucl\b"]),
    ("coupe",  [r"coupe de france", r"europa league", r"conference league",
                r"coupe d'europe", r"trophée des champions", r"supercoupe"]),
    ("edf",    [r"équipe de france", r"equipe de france", r"\bbleus\b",
                r"éliminatoires", r"eliminatoires", r"nations league"]),
    ("padel",  [r"padel"]),
    ("f1",     [r"formule 1", r"formula 1", r"\bf1\b"]),
    ("rugby",  [r"rugby", r"top 14", r"champions cup", r"xv de france", r"six nations"]),
    ("tennis", [r"tennis", r"roland[- ]garros", r"wimbledon", r"open d'australie",
                r"\bus open\b", r"\batp\b", r"\bwta\b", r"coupe davis"]),
    ("foot",   [r"football", r"\bfoot\b", r"ligue 2", r"premier league",
                r"bundesliga", r"serie a", r"\bliga\b"]),
    ("sport",  [r"basket", r"handball", r"\bhand\b", r"volley", r"cyclisme",
                r"athlétisme", r"natation", r"jeux olympiques", r"\bjo\b",
                r"boxe", r"\bmma\b", r"golf", r"ski", r"biathlon", r"moto",
                r"rallye", r"nba\b", r"nfl\b", r"\bufc\b"]),
]
# Mots-clés indiquant un genre sport (dans les balises <category> du flux)
SPORT_GENRE_KEYS = ["sport", "football", "rugby", "tennis", "basket", "hand",
                    "volley", "cyclisme", "athlétisme", "athletisme", "match",
                    "compétition", "competition", "formule"]

# Catégories considérées comme « sport »
SPORT_CATEGORIES = {"ldc", "coupe", "edf", "foot", "padel", "tennis",
                    "rugby", "f1", "sport"}

# Chaînes généralistes (TNT) : on ne garde QUE leurs programmes sport.
# Les autres chaînes du centre (Canal+, Canal+ Sport/Foot, beIN, L'Équipe)
# gardent TOUS leurs programmes.
SPORT_ONLY_CHANNELS = {"TF1", "France 2", "France 3", "M6", "W9", "TMC"}

# Titres « vides » / remplissage à ignorer (surtout Canal+ Live numérotés) :
# bandeaux d'attente, annonces « À venir : … », marqueurs « Terminé : … ».
FILLER_TITLES = [
    r"vivez en direct",
    r"^\s*(a|à)\W*venir\b",
    r"^\s*termin",
    r"^\s*fin des programmes",
    r"programmes de la nuit",
]

# ── Genres XMLTV (non-sport) → catégorie. Ordre important. ────────────────────
GENRE_MAP = [
    ("jeunesse", ["jeunesse", "enfant", "dessin animé", "dessin anime",
                  "animation", "manga", "anime"]),
    ("film",     ["téléfilm", "telefilm", "film", "cinéma", "cinema",
                  "long métrage", "long metrage", "court métrage"]),
    ("serie",    ["série", "serie", "feuilleton", "soap", "saga", "sitcom"]),
    ("info",     ["information", "journal", "météo", "meteo", "politique",
                  "débat", "debat", "actualité", "actualite", "\binfo\b"]),
    ("doc",      ["documentaire", "reportage", "découverte", "decouverte",
                  "nature", "animalier", "histoire", "société", "societe",
                  "science"]),
    ("culture",  ["musique", "concert", "théâtre", "theatre", "opéra", "opera",
                  "spectacle", "\bart\b", "culture", "danse"]),
    ("mag",      ["magazine", "divertissement", "talk", "\bjeu\b", "téléréalité",
                  "telerealite", "variété", "variete", "humour", "cuisine",
                  "téléachat", "teleachat", "émission", "emission"]),
]


def log(msg):
    print(msg, file=sys.stderr)


def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def read_source(source):
    """Retourne le contenu XMLTV (str) depuis une URL ou un fichier, gzip géré."""
    if re.match(r"^https?://", source):
        log(f"Téléchargement : {source}")
        req = urllib.request.Request(source, headers={"User-Agent": "urbansoccer-tv/1.0"})
        with urllib.request.urlopen(req, timeout=60) as r:  # proxies via env
            raw = r.read()
    else:
        log(f"Lecture fichier : {source}")
        with open(source, "rb") as f:
            raw = f.read()
    # gzip ? (magic 1f 8b, ou extension)
    if raw[:2] == b"\x1f\x8b" or source.endswith(".gz"):
        raw = gzip.GzipFile(fileobj=io.BytesIO(raw)).read()
    return raw.decode("utf-8", errors="replace")


def parse_xmltv_time(s):
    """'20260923201500 +0200' → datetime naïf (heure locale du flux)."""
    if not s:
        return None
    m = re.match(r"\s*(\d{14})", s)
    if not m:
        return None
    return datetime.strptime(m.group(1), "%Y%m%d%H%M%S")


def guess_category(title, subtitle, genres):
    """Classe un programme (sport affiné, sinon genre, sinon 'autre')."""
    text = f"{title} {subtitle}".lower()
    genre_str = " ".join(genres).lower()

    is_sport = any(k in genre_str for k in SPORT_GENRE_KEYS) or \
        any(re.search(p, text) for _, pats in SPORT_RULES for p in pats)
    if is_sport:
        for cat, patterns in SPORT_RULES:
            if any(re.search(p, text) for p in patterns):
                return cat
        return "sport"

    for cat, keys in GENRE_MAP:
        if any(re.search(k, genre_str) for k in keys):
            return cat
    return "autre"


def iter_programmes(xml_text):
    """Yield (channel_id, start, stop, title, subtitle, genres[]) depuis le XMLTV."""
    import xml.etree.ElementTree as ET
    for _, elem in ET.iterparse(io.StringIO(xml_text), events=("end",)):
        if elem.tag != "programme":
            continue
        channel = elem.get("channel", "")
        start = parse_xmltv_time(elem.get("start", ""))
        stop = parse_xmltv_time(elem.get("stop", ""))
        title = (elem.findtext("title") or "").strip()
        subtitle = (elem.findtext("sub-title") or "").strip()
        genres = [(c.text or "").strip().lower() for c in elem.findall("category")]
        yield channel, start, stop, title, subtitle, genres
        elem.clear()


def main():
    ap = argparse.ArgumentParser(description="Pré-remplit un brouillon de programme TV depuis un flux XMLTV.")
    ap.add_argument("--source", required=True, help="URL ou fichier XMLTV (.xml ou .xml.gz)")
    ap.add_argument("--days", type=int, default=7, help="Nombre de jours à partir d'aujourd'hui (défaut 7)")
    ap.add_argument("--out", default=DEFAULT_OUT, help="Fichier brouillon de sortie")
    args = ap.parse_args()

    prog = load_json(PROG_PATH)
    dispo = set(prog.get("abonnement", {}).get("disponibles", []))
    if not os.path.exists(MAPPING_PATH):
        log(f"⚠️  Mapping introuvable : {MAPPING_PATH}")
        log("   Créez-le : {\"<id XMLTV>\": \"<nom de chaîne du centre>\", ...}")
        sys.exit(1)
    mapping = load_json(MAPPING_PATH)
    # on ne garde du mapping que les chaînes réellement dans l'abonnement
    mapping = {xid: name for xid, name in mapping.items() if name in dispo}
    if not mapping:
        log("⚠️  Aucune chaîne du mapping ne figure dans abonnement.disponibles.")
        sys.exit(1)

    today = datetime.now().date()
    date_min = today
    date_max = today + timedelta(days=args.days)

    xml_text = read_source(args.source)

    kept = []
    seen_channels = set()
    for channel, start, stop, title, subtitle, genres in iter_programmes(xml_text):
        if channel not in mapping:
            continue
        seen_channels.add(channel)
        if not start:
            continue
        if not (date_min <= start.date() < date_max):
            continue

        cat = guess_category(title, subtitle, genres)

        # Programmes « vides » / remplissage (Canal+ Live en attente…) → ignorés.
        _t = (title or subtitle or "").lower()
        if any(re.search(p, _t) for p in FILLER_TITLES):
            continue

        # Chaînes généralistes (TNT) : on ne garde que le sport.
        if mapping[channel] in SPORT_ONLY_CHANNELS and cat not in SPORT_CATEGORIES:
            continue

        duree = 60
        if stop and stop > start:
            duree = int((stop - start).total_seconds() // 60)

        # On retire les petits programmes (< 30 min) : flashs, pastilles, bandes-annonces…
        if duree < 30:
            continue

        # Horaires d'ouverture du centre : 9h → minuit en semaine, 9h → 23h le week-end.
        # On ne garde que ce qui chevauche ce créneau.
        close_h = 23 if start.weekday() >= 5 else 24
        sh = start.hour + start.minute / 60
        eh = sh + duree / 60
        if not (sh < close_h and eh > 9):
            continue

        kept.append({
            "date": start.strftime("%Y-%m-%d"),
            "heure": start.strftime("%H:%M"),
            "categorie": cat,
            "competition": title or subtitle,
            "affiche": subtitle or title,
            "chaine": mapping[channel],
            "duree_min": duree,
        })

    kept.sort(key=lambda e: (e["date"], e["heure"], e["chaine"]))

    out = {
        "_meta": {
            **prog.get("_meta", {}),
            "updated_at": datetime.now().strftime("%Y-%m-%d"),
            "source": f"XMLTV — {args.source}",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "note": "Généré automatiquement depuis un flux XMLTV. Éditable à la main si besoin.",
        },
        "abonnement": prog.get("abonnement", {}),
        "chaines_meta": prog.get("chaines_meta", {}),
        "categories": CATEGORIES,
        "evenements": kept,
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
        f.write("\n")

    # ── Résumé ──
    by_chan = {}
    for e in kept:
        by_chan[e["chaine"]] = by_chan.get(e["chaine"], 0) + 1
    log("")
    log(f"Chaînes du mapping trouvées dans le flux : {len(seen_channels)}/{len(mapping)}")
    for xid, name in mapping.items():
        flag = "✓" if xid in seen_channels else "✗ (absente du flux)"
        log(f"  {flag}  {xid} → {name}")
    log("")
    log(f"{len(kept)} programmes retenus sur {args.days} j :")
    for name, n in sorted(by_chan.items()):
        log(f"  {n:3d}  {name}")
    log("")
    log(f"Brouillon écrit : {args.out}")
    log("→ Relisez, retirez ce qui n'a pas d'intérêt, corrigez les libellés,")
    log("  renommez en data/tv-programme.json, puis lancez : python generate_tv.py")


if __name__ == "__main__":
    main()

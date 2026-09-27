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
  3. ne garde que les programmes « sport » pertinents (genre ou mots-clés) ;
  4. devine la catégorie (ldc / coupe / edf / foot / padel / tennis / rugby / f1…) ;
  5. écrit un brouillon au schéma de data/tv-programme.json.

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

# ── Détection sport + catégorie à partir du titre / genre XMLTV ──────────────
# Ordre important : la première règle qui matche gagne.
CATEGORY_RULES = [
    ("ldc",    [r"ligue des champions", r"champions league", r"c1\b", r"ucl\b"]),
    ("coupe",  [r"coupe de france", r"europa league", r"conference league",
                r"coupe d'europe", r"trophée des champions", r"supercoupe"]),
    ("edf",    [r"équipe de france", r"equipe de france", r"bleus",
                r"éliminatoires", r"eliminatoires", r"nations league"]),
    ("padel",  [r"padel", r"premier padel"]),
    ("f1",     [r"formule 1", r"formula 1", r"\bf1\b", r"grand prix", r"\bgp\b"]),
    ("rugby",  [r"rugby", r"top 14", r"champions cup", r"xv de france", r"six nations",
                r"tournoi des (6|six) nations"]),
    ("tennis", [r"tennis", r"roland[- ]garros", r"wimbledon", r"open d'australie",
                r"us open", r"atp\b", r"wta\b", r"coupe davis"]),
    ("foot",   [r"football", r"foot\b", r"ligue 2", r"premier league",
                r"bundesliga", r"serie a", r"liga\b", r"match", r"j\d{1,2}\b"]),
    ("sport",  [r"basket", r"handball", r"volley", r"cyclisme", r"athlétisme",
                r"natation", r"jeux olympiques", r"boxe", r"mma", r"golf"]),
]

# Genres XMLTV considérés comme « sport »
SPORT_GENRES = {"sport", "sports", "sporting event", "football", "rugby",
                "tennis", "basketball", "match", "compétition"}

# Ne jamais proposer (règle éditoriale du centre)
EXCLUDE_KEYWORDS = [r"ligue 1\b", r"\bl1\b"]


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


def guess_category(text):
    low = text.lower()
    for cat, patterns in CATEGORY_RULES:
        for p in patterns:
            if re.search(p, low):
                return cat
    return None


def is_excluded(text):
    low = text.lower()
    return any(re.search(p, low) for p in EXCLUDE_KEYWORDS)


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

        text = f"{title} {subtitle}"
        if is_excluded(text):
            continue

        is_sport_genre = any(g in SPORT_GENRES for g in genres)
        cat = guess_category(text)
        if not is_sport_genre and not cat:
            continue  # ni genre sport, ni mot-clé reconnu → ignoré
        if not cat:
            cat = "sport"

        duree = 130
        if stop and stop > start:
            duree = int((stop - start).total_seconds() // 60)

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
            "titre": prog.get("_meta", {}).get("titre", "Programme TV"),
            "source": f"XMLTV — {args.source}",
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "note": "BROUILLON à relire/élaguer, puis renommer en tv-programme.json.",
        },
        "abonnement": prog.get("abonnement", {}),
        "chaines_meta": prog.get("chaines_meta", {}),
        "categories": prog.get("categories", {}),
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

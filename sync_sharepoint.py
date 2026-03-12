#!/usr/bin/env python3
"""
Synchronisation SharePoint → GitHub des fichiers de plannings.

Ce script télécharge les fichiers "Plannings YYYY SXX.xlsx" depuis un dossier
SharePoint (via l'API Microsoft Graph) et les commit/push vers le repo GitHub.

Prérequis :
    1. Créer une App Registration dans Azure AD (voir README-SYNC.md)
    2. Définir les variables d'environnement ou les secrets GitHub :
        - AZURE_TENANT_ID
        - AZURE_CLIENT_ID
        - AZURE_CLIENT_SECRET
        - SHAREPOINT_SITE_ID    (ou SHAREPOINT_SITE_URL)
        - SHAREPOINT_FOLDER_PATH (ex: "General/RH/Plannings")

Usage :
    python sync_sharepoint.py                    # sync tous les plannings
    python sync_sharepoint.py --dry-run          # voir les fichiers sans télécharger
    python sync_sharepoint.py --pattern "S11"    # sync un seul fichier
"""

import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime
from urllib.parse import quote

try:
    import requests
except ImportError:
    print("Installation de 'requests'...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "requests"])
    import requests


# ── Configuration ────────────────────────────────────────────────────────────

AZURE_TENANT_ID = os.environ.get("AZURE_TENANT_ID", "")
AZURE_CLIENT_ID = os.environ.get("AZURE_CLIENT_ID", "")
AZURE_CLIENT_SECRET = os.environ.get("AZURE_CLIENT_SECRET", "")

SHAREPOINT_SITE_URL = os.environ.get("SHAREPOINT_SITE_URL", "")
SHAREPOINT_SITE_ID = os.environ.get("SHAREPOINT_SITE_ID", "")
SHAREPOINT_DRIVE_ID = os.environ.get("SHAREPOINT_DRIVE_ID", "")
SHAREPOINT_FOLDER_PATH = os.environ.get(
    "SHAREPOINT_FOLDER_PATH", "General/RH/Plannings"
)

# Pattern pour les fichiers de planning
PLANNING_PATTERN = re.compile(
    r"^Plannings\s+\d{4}\s+S\d{1,2}(?:\s+v\d+)?\.xlsx$", re.IGNORECASE
)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


# ── Authentification Microsoft Graph ─────────────────────────────────────────

def get_access_token() -> str:
    """Obtient un token d'accès via client credentials (app-only)."""
    url = f"https://login.microsoftonline.com/{AZURE_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": AZURE_CLIENT_ID,
        "client_secret": AZURE_CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default",
    }
    resp = requests.post(url, data=data, timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]


def graph_headers(token: str) -> dict:
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


# ── Résolution du site et du drive SharePoint ────────────────────────────────

def resolve_site_id(token: str) -> str:
    """Résout l'ID du site SharePoint depuis l'URL."""
    if SHAREPOINT_SITE_ID:
        return SHAREPOINT_SITE_ID

    # URL format: https://contoso.sharepoint.com/sites/MySite
    # Graph API: /sites/{hostname}:/{server-relative-path}
    from urllib.parse import urlparse
    parsed = urlparse(SHAREPOINT_SITE_URL)
    hostname = parsed.hostname
    site_path = parsed.path.rstrip("/")

    url = f"{GRAPH_BASE}/sites/{hostname}:{site_path}"
    resp = requests.get(url, headers=graph_headers(token), timeout=30)
    resp.raise_for_status()
    site_id = resp.json()["id"]
    print(f"  Site ID résolu : {site_id}")
    return site_id


def resolve_drive_id(token: str, site_id: str) -> str:
    """Résout l'ID du drive (bibliothèque Documents) du site."""
    if SHAREPOINT_DRIVE_ID:
        return SHAREPOINT_DRIVE_ID

    url = f"{GRAPH_BASE}/sites/{site_id}/drives"
    resp = requests.get(url, headers=graph_headers(token), timeout=30)
    resp.raise_for_status()
    drives = resp.json().get("value", [])

    # Cherche le drive "Documents" (par défaut)
    for d in drives:
        if d.get("name", "").lower() in ("documents", "documents partagés"):
            print(f"  Drive ID résolu : {d['id']} ({d['name']})")
            return d["id"]

    # Fallback : premier drive
    if drives:
        print(f"  Drive ID (fallback) : {drives[0]['id']} ({drives[0]['name']})")
        return drives[0]["id"]

    raise RuntimeError("Aucun drive trouvé sur le site SharePoint")


# ── Liste et téléchargement des fichiers ─────────────────────────────────────

def list_planning_files(token: str, drive_id: str, pattern: str = None) -> list:
    """Liste les fichiers de planning dans le dossier SharePoint."""
    folder_path = quote(SHAREPOINT_FOLDER_PATH)
    url = f"{GRAPH_BASE}/drives/{drive_id}/root:/{folder_path}:/children"

    params = {"$select": "name,id,size,lastModifiedDateTime,file"}
    resp = requests.get(url, headers=graph_headers(token), params=params, timeout=30)
    resp.raise_for_status()

    files = []
    for item in resp.json().get("value", []):
        name = item.get("name", "")
        # Ignorer les fichiers temporaires (~$...)
        if name.startswith("~$"):
            continue
        # Ne garder que les fichiers Excel de planning
        if not PLANNING_PATTERN.match(name):
            continue
        # Filtrer par pattern si spécifié
        if pattern and pattern.lower() not in name.lower():
            continue
        files.append(item)

    return sorted(files, key=lambda f: f["name"])


def download_file(token: str, drive_id: str, item_id: str, dest_path: str):
    """Télécharge un fichier depuis SharePoint."""
    url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
    resp = requests.get(url, headers=graph_headers(token), timeout=120)
    resp.raise_for_status()

    with open(dest_path, "wb") as f:
        f.write(resp.content)


# ── Git operations ───────────────────────────────────────────────────────────

def git_has_changes() -> bool:
    """Vérifie s'il y a des changements dans le repo."""
    result = subprocess.run(
        ["git", "status", "--porcelain"],
        capture_output=True, text=True, cwd=SCRIPT_DIR
    )
    return bool(result.stdout.strip())


def git_commit_and_push(files: list[str], message: str):
    """Commit et push les fichiers modifiés."""
    for f in files:
        subprocess.run(["git", "add", f], cwd=SCRIPT_DIR, check=True)

    subprocess.run(
        ["git", "commit", "-m", message],
        cwd=SCRIPT_DIR, check=True
    )
    subprocess.run(
        ["git", "push"],
        cwd=SCRIPT_DIR, check=True
    )


# ── Synchronisation principale ───────────────────────────────────────────────

def sync(dry_run: bool = False, pattern: str = None, no_push: bool = False):
    """Synchronise les plannings de SharePoint vers le repo local."""
    print("=" * 60)
    print("  Sync SharePoint → GitHub : Plannings UrbanSoccer")
    print("=" * 60)

    # Vérification des credentials
    missing = []
    if not AZURE_TENANT_ID:
        missing.append("AZURE_TENANT_ID")
    if not AZURE_CLIENT_ID:
        missing.append("AZURE_CLIENT_ID")
    if not AZURE_CLIENT_SECRET:
        missing.append("AZURE_CLIENT_SECRET")
    if not SHAREPOINT_SITE_URL and not SHAREPOINT_SITE_ID:
        missing.append("SHAREPOINT_SITE_URL ou SHAREPOINT_SITE_ID")

    if missing:
        print("\n❌ Variables d'environnement manquantes :")
        for m in missing:
            print(f"   - {m}")
        print("\nVoir README-SYNC.md pour la configuration.")
        sys.exit(1)

    # Auth
    print("\n🔐 Authentification Microsoft Graph...")
    token = get_access_token()
    print("  ✅ Token obtenu")

    # Résolution site + drive
    print("\n📂 Résolution du site SharePoint...")
    site_id = resolve_site_id(token)
    drive_id = resolve_drive_id(token, site_id)

    # Liste des fichiers
    print(f"\n📋 Fichiers dans '{SHAREPOINT_FOLDER_PATH}' :")
    files = list_planning_files(token, drive_id, pattern)

    if not files:
        print("  Aucun fichier de planning trouvé.")
        return

    for f in files:
        size_kb = f.get("size", 0) / 1024
        modified = f.get("lastModifiedDateTime", "?")[:10]
        print(f"  📄 {f['name']}  ({size_kb:.0f} Ko, modifié {modified})")

    if dry_run:
        print(f"\n🔍 Mode dry-run : {len(files)} fichier(s) trouvé(s), rien téléchargé.")
        return

    # Téléchargement
    print(f"\n⬇️  Téléchargement de {len(files)} fichier(s)...")
    downloaded = []
    for f in files:
        dest = os.path.join(SCRIPT_DIR, f["name"])
        print(f"  ⬇️  {f['name']}...", end=" ", flush=True)
        download_file(token, drive_id, f["id"], dest)
        downloaded.append(f["name"])
        print("✅")

    # Git commit + push
    if not git_has_changes():
        print("\n✅ Aucun changement détecté (fichiers identiques).")
        return

    week_nums = []
    for name in downloaded:
        m = re.search(r"S(\d{1,2})", name)
        if m:
            week_nums.append(f"S{m.group(1)}")
    weeks_str = ", ".join(sorted(set(week_nums)))

    message = f"📅 Sync plannings SharePoint : {weeks_str or ', '.join(downloaded)}"
    print(f"\n📤 Commit : {message}")

    if no_push:
        # Just stage the files, don't commit
        for f in downloaded:
            subprocess.run(["git", "add", f], cwd=SCRIPT_DIR)
        print("  Fichiers stagés (--no-push). Committez manuellement.")
    else:
        git_commit_and_push(downloaded, message)
        print("  ✅ Push réussi !")

    print(f"\n🎉 Synchronisation terminée : {len(downloaded)} fichier(s)")


# ── CLI ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Synchronise les plannings depuis SharePoint vers GitHub"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Lister les fichiers sans les télécharger"
    )
    parser.add_argument(
        "--pattern", type=str, default=None,
        help="Filtrer par nom (ex: 'S11' pour ne sync que la semaine 11)"
    )
    parser.add_argument(
        "--no-push", action="store_true",
        help="Stage les fichiers sans commit/push"
    )
    args = parser.parse_args()
    sync(dry_run=args.dry_run, pattern=args.pattern, no_push=args.no_push)


if __name__ == "__main__":
    main()

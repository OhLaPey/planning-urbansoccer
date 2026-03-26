#!/usr/bin/env python3
"""
Sync retour : GitHub (JSON) → Excel → SharePoint.

Quand un utilisateur coche des présences sur la page web, le JavaScript
met à jour le JSON sur GitHub. Ce script :
  1. Télécharge le fichier Excel de présences depuis SharePoint
  2. Applique les différences entre les JSON GitHub et l'Excel
  3. Re-uploade l'Excel modifié vers SharePoint

Usage :
    python sync_presences_to_sharepoint.py             # sync complet
    python sync_presences_to_sharepoint.py --dry-run    # voir les changements sans appliquer
"""

import argparse
import os
import sys
import tempfile

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def main():
    parser = argparse.ArgumentParser(
        description="Sync présences : JSON GitHub → Excel → SharePoint"
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Afficher les changements sans les appliquer"
    )
    args = parser.parse_args()

    # Import sync modules
    try:
        from sync_sharepoint import (
            get_access_token, resolve_site_id, resolve_drive_id,
            find_attendance_item, download_file, upload_file,
            SHAREPOINT_FOLDER_PATH, AZURE_TENANT_ID, AZURE_CLIENT_ID,
            AZURE_CLIENT_SECRET, SHAREPOINT_SITE_URL, SHAREPOINT_SITE_ID,
        )
    except ImportError as e:
        print(f"Erreur d'import : {e}")
        sys.exit(1)

    from sync_presences import sync as sync_json_to_excel
    from attendance import find_attendance_file, parse_attendance

    print("=" * 60)
    print("  Sync retour : Page web → JSON → Excel → SharePoint")
    print("=" * 60)

    # ── 1. Check credentials ──
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
        print(f"\nVariables manquantes : {', '.join(missing)}")
        print("Voir README-SYNC.md pour la configuration.")
        sys.exit(1)

    # ── 2. Auth + resolve SharePoint ──
    print("\nAuthentification Microsoft Graph...")
    token = get_access_token()
    site_id = resolve_site_id(token)
    drive_id = resolve_drive_id(token, site_id)

    # ── 3. Find attendance file on SharePoint ──
    print("\nRecherche du fichier de présences sur SharePoint...")
    sp_item = find_attendance_item(token, drive_id)
    if not sp_item:
        print("Aucun fichier de présences trouvé sur SharePoint.")
        sys.exit(1)

    sp_name = sp_item["name"]
    sp_id = sp_item["id"]
    print(f"  Trouvé : {sp_name}")

    # ── 4. Download to local ──
    local_path = os.path.join(SCRIPT_DIR, sp_name)
    print(f"\nTéléchargement de {sp_name}...")
    download_file(token, drive_id, sp_id, local_path)
    print(f"  Enregistré : {local_path}")

    # ── 5. Parse current state ──
    data_before = parse_attendance(local_path)
    before_summary = {}
    for c in data_before["creneaux"]:
        total = sum(
            sum(1 for v in k["attendance"].values() if v == 1)
            for g in c["groups"] for k in g["kids"]
        )
        before_summary[c["slug"]] = total

    # ── 6. Apply JSON changes to Excel ──
    print("\nApplication des modifications JSON → Excel...")
    if args.dry_run:
        print("  (mode dry-run, aucune modification appliquée)")
    else:
        sync_json_to_excel()

    # ── 7. Check if anything changed ──
    data_after = parse_attendance(local_path)
    changed = False
    for c in data_after["creneaux"]:
        total = sum(
            sum(1 for v in k["attendance"].values() if v == 1)
            for g in c["groups"] for k in g["kids"]
        )
        if before_summary.get(c["slug"]) != total:
            changed = True
            diff = total - before_summary.get(c["slug"], 0)
            sign = "+" if diff > 0 else ""
            print(f"  {c['title']} : {sign}{diff} présences")

    if not changed:
        print("\nAucun changement à remonter vers SharePoint.")
        return

    if args.dry_run:
        print("\n(dry-run) Les changements ci-dessus seraient uploadés vers SharePoint.")
        return

    # ── 8. Upload modified Excel back to SharePoint ──
    print(f"\nUpload de {sp_name} vers SharePoint...")
    upload_file(token, drive_id, SHAREPOINT_FOLDER_PATH, sp_name, local_path)
    print("  Upload réussi !")

    print("\nSync terminée : page web → Excel SharePoint")


if __name__ == "__main__":
    main()

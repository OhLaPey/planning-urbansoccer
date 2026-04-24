# Planning Urban Soccer

## Repo

- Auto-merge activé : utiliser `gh pr merge --auto -m` après création de PR
- Suppression automatique des branches après merge
- Ne JAMAIS push directement sur main (branche protégée)
- Ne JAMAIS inclure les fichiers data/SXX-events.json modifiés par le web dans les PR (garder les versions de main)
- Restaurer les notes/ et ics/ depuis main avant de commiter

## Stack

- Python (generate.py), GitHub Pages (main), fichiers Excel source
- HTML avec `var DATA = {...}` embarqué (pas d'API, tout est statique)
- Workflow generate.yml : se déclenche à chaque push sur main, regénère les HTML et déploie sur GitHub Pages

## generate.py — Points critiques

- `generate_attendance_pages()` : NE PAS appeler (les pages présences sont gérées séparément)
- Les modifs web (source "Modif admin" dans les JSON) doivent être préservées : generate.py lit le JSON existant avant de l'écraser
- Le fix chevauchement DÉCOUPE les events au lieu de tronquer (VDC autour d'un L-ARB)
- Le workflow sauvegarde les presences*.html AVANT generate.py et les restaure après

## Fichiers à ne PAS écraser

- `data/SXX-events.json` quand source = "Modif admin" (modifs web)
- `notes/SXX.json` (notes + remplacements ajoutés via le web)
- `presences-*.html` (pages maintenues avec features : checkboxes, navigation, highlight)

## Sync

- **Plannings Excel → Web** : Gmail Apps Script → GitHub → workflow generate.yml
- **Plannings Web → Excel** : mode édition → JSON poussé sur GitHub → generate.py préserve les modifs
- **Présences Excel → Web** : Power Automate → email → Apps Script → GitHub
- **Présences Web → Excel** : Apps Script surveille GitHub → email → Power Automate → script Office

## Commandes utiles

- `python generate.py` — régénère tous les HTML/JSON/ICS
- `gh pr create --auto-merge` — créer et auto-merger une PR

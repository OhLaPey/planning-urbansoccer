# Fix Planning Staff — Pages non affichées

## Contexte

Site de planning d'équipe déployé sur GitHub Pages (branche main). Le système fonctionne ainsi : des fichiers Excel (par semaine) sont synchronisés depuis SharePoint → generate.py les lit et génère des pages HTML + JSON + ICS → le tout est poussé sur main → GitHub Pages les affiche.

Depuis la dernière PR (#51, branche claude/calendar-subscription-setup-FxCut), les plannings ne s'affichent plus sur l'interface web. La PR ajoutait un champ _meta dans les events.json et une note discrète en bas de page. Des conflits de merge sur data/S12-events.json et data/S13-events.json ont été résolus puis un force push a été fait, mais la session Claude Code a crashé après. On ne sait pas si le merge dans main a abouti.

## Tâche 1 — Diagnostiquer et fixer

1. Vérifier l'état de la PR #51 : est-elle mergée dans main ? Si non, la merger
2. Vérifier que les fichiers HTML existent sur main pour les semaines actuelles (S12, S13, S14)
3. Vérifier que index.html redirige vers la bonne semaine (on est en S13)
4. Vérifier que les fichiers data/SXX-events.json sont valides (pas de conflits git résiduels <<<<<<<)
5. Vérifier que generate.py tourne sans erreur sur tous les Excel disponibles
6. Si des fichiers manquent sur main : relancer generate.py, commit et push sur main
7. Vérifier le déploiement GitHub Pages (settings du repo, branche source)

## Commandes utiles

* `python generate.py` — régénère tous les HTML/JSON/ICS
* `gh pr list` — voir les PR ouvertes
* `gh pr merge 51` — merger si nécessaire
* `git log --oneline -10` — derniers commits sur la branche courante

## Stack

* Python (generate.py), GitHub Pages (main), fichiers Excel source
* HTML avec `var DATA = {...}` embarqué (pas d'API, tout est statique)

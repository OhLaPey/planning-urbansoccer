# Programme TV du centre

Page à diffuser sur les écrans du centre : elle affiche les **événements sportifs
majeurs de la semaine** (Ligue des Champions, Coupes, Équipe de France, Padel,
Top 14…) **uniquement sur les chaînes dont dispose le centre**, avec le **numéro
de chaîne** pour zapper vite.

Même principe que les plannings :

```
data/tv-programme.json   →   python generate_tv.py   →   tv.html   →   GitHub Pages
```

La page est **statique** (aucune API). Tout le rendu « du jour / en direct » se
fait dans le navigateur, donc elle reste juste toute la journée sans régénérer.

## Voir la page

Une fois publiée sur GitHub Pages : **`.../tv.html`**
(à afficher en plein écran sur les TV — la page se recharge seule 1×/h et
rafraîchit l'état « EN DIRECT » chaque minute).

## Mettre à jour le programme (chaque semaine)

1. Ouvrir **`data/tv-programme.json`**
2. Dans `evenements`, remplacer les exemples par les vrais matchs/événements :

   ```json
   {
     "date": "2026-09-16",
     "heure": "21:00",
     "categorie": "ldc",
     "competition": "Ligue des Champions — J1",
     "affiche": "PSG — Atalanta",
     "chaine": "Canal+"
   }
   ```

3. Lancer :

   ```bash
   python generate_tv.py
   ```

4. Commit + push sur `main` → GitHub Pages met la page à jour.

## Règle de diffusion

- Un événement n'est proposé à la diffusion **que si sa `chaine` figure dans
  `abonnement.disponibles`** (Canal+, beIN Sports, chaînes en clair).
- Les événements sur `abonnement.non_disponibles` (**Ligue 1+**, **DAZN**) sont
  affichés à part, en grisé, dans « ⛔ Non disponible au centre » — pour que
  l'équipe sache pourquoi un gros match n'est pas diffusable.

## Champs d'un événement

| Champ         | Obligatoire | Exemple                          |
|---------------|-------------|----------------------------------|
| `date`        | oui         | `"2026-09-16"` (AAAA-MM-JJ)      |
| `heure`       | oui         | `"21:00"`                        |
| `categorie`   | oui         | `ldc`, `coupe`, `edf`, `foot`, `padel`, `tennis`, `rugby`, `sport` |
| `competition` | oui         | `"Ligue des Champions — J1"`     |
| `affiche`     | oui         | `"PSG — Atalanta"`               |
| `chaine`      | oui         | doit exister dans `chaines_meta` |

## Chaînes et numéros

Les numéros sont dans `chaines_meta` (`"numero"`).

> ⚠️ **À vérifier sur le décodeur/box du centre.** Les numéros Canal+ et beIN
> Sports dépendent de l'opérateur (Canal, Free, Orange, SFR, Bouygues). Les
> chaînes en clair (TNT) sont nationales et fiables : TF1 = 1, France 2 = 2,
> France 3 = 3, M6 = 6, W9 = 9, TMC = 10, L'Équipe = 21.

Pour ajouter une chaîne : l'ajouter dans `chaines_meta` (numéro + couleur +
`clair`) puis, si elle est captée au centre, dans `abonnement.disponibles`.

## Repère : quelle compétition sur quelle chaîne (France)

Aide-mémoire pour remplir `chaine` (à ajuster selon les droits de la saison) :

| Compétition                     | Chaîne(s) au centre                 |
|---------------------------------|-------------------------------------|
| Ligue des Champions             | Canal+ / beIN Sports                |
| Premier League, Liga, Serie A   | Canal+ (droits Canal)               |
| Coupe de France                 | beIN Sports + France TV (en clair)  |
| Équipe de France                | TF1 / M6 (en clair)                 |
| Premier Padel                   | Canal+ (Canal+ Sport)               |
| Top 14 / rugby                  | Canal+                              |
| **Ligue 1**                     | **Ligue 1+ / DAZN → non diffusable**|

*(La Ligue 1 est majoritairement sur Ligue 1+ / DAZN, hors abonnement : un match
Ligue 1 diffusé par beIN Sports reste, lui, diffusable.)*

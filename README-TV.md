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

Un **sélecteur de jour** (même principe que le planning staff) permet de
naviguer d'un jour à l'autre ; par défaut, le jour du jour est affiché. Seuls
les jours ayant au moins une diffusion apparaissent dans le sélecteur.

Chaque journée est présentée en **guide TV horizontal** (comme le planning) :
- à **gauche**, les **chaînes pertinentes ce jour-là** (uniquement celles qui
  diffusent quelque chose), avec leur numéro, **triées par numéro croissant** ;
- au **centre**, le **détail des diffusions** de chaque chaîne sur l'axe des
  heures, **découpé par programme** (plusieurs barres sur une même ligne si la
  chaîne enchaîne plusieurs diffusions) ;
- barres colorées par catégorie ; le jour même, un **curseur d'heure en direct**
  (trait doré) suit l'heure réelle et se recentre automatiquement, et les
  diffusions en cours portent un badge « Direct ».

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

- Un événement n'est affiché **que si sa `chaine` figure dans
  `abonnement.disponibles`** (Canal+, beIN Sports, chaînes en clair).
- Les événements sur `abonnement.non_disponibles` (**Ligue 1+**, **DAZN**,
  **Disney+**) ne sont **pas affichés** — la page ne montre que ce qui est
  réellement diffusable au centre. (Vous pouvez tout de même les laisser dans le
  fichier de données : ils seront simplement ignorés à l'affichage.)
- **La Ligue 1 est 100 % sur Ligue 1+** (hors abonnement) : un match de L1
  n'apparaît donc jamais comme diffusable. beIN Sports n'a **plus** la Ligue 1
  ni la **Liga** (partie sur DAZN + Disney+ en 2026), mais conserve la
  **Ligue 2** (diffusable).
- Garde-fou : pour retirer manuellement un événement de la diffusion **même s'il
  passe sur une chaîne disponible**, ajouter `"diffusable": false` (+ `"raison"`).
  Il bascule alors dans la section grisée.
- **Rugby (Top 14, sur Canal), Formule 1 (Canal+)** : diffusables, à ajouter
  comme n'importe quel événement.

## Champs d'un événement

| Champ         | Obligatoire | Exemple                          |
|---------------|-------------|----------------------------------|
| `date`        | oui         | `"2026-09-16"` (AAAA-MM-JJ)      |
| `heure`       | oui         | `"21:00"`                        |
| `categorie`   | oui         | `ldc`, `coupe`, `edf`, `foot`, `padel`, `tennis`, `rugby`, `sport` |
| `competition` | oui         | `"Ligue des Champions — J1"`     |
| `affiche`     | oui         | `"PSG — Atalanta"`               |
| `chaine`      | oui         | doit exister dans `chaines_meta` |
| `duree_min`   | non         | durée de la barre en minutes (déf. 130) |
| `diffusable`  | non         | `false` pour forcer le masquage  |

## Chaînes et numéros

Les numéros sont dans `chaines_meta` (`"numero"`).

> ⚠️ **À vérifier sur le décodeur/box du centre.** Les numéros Canal+ et beIN
> Sports dépendent de l'opérateur (Canal, Free, Orange, SFR, Bouygues). Les
> chaînes en clair (TNT) sont nationales et fiables : TF1 = 1, France 2 = 2,
> France 3 = 3, M6 = 6, W9 = 9, TMC = 10, L'Équipe = 21.

Pour ajouter une chaîne : l'ajouter dans `chaines_meta` (numéro + couleur +
`clair`) puis, si elle est captée au centre, dans `abonnement.disponibles`.

## Repère : quelle compétition sur quelle chaîne (France)

Aide-mémoire pour remplir `chaine` (droits saison **2026-2027** — à revérifier
chaque saison) :

| Compétition                     | Chaîne(s)              | Au centre ?          |
|---------------------------------|------------------------|----------------------|
| Ligue des Champions             | Canal+                 | ✅ diffusable        |
| Premier League, Europa League   | Canal+                 | ✅ diffusable        |
| **Ligue 2**                     | **beIN Sports**        | ✅ diffusable        |
| Bundesliga                      | beIN Sports            | ✅ diffusable        |
| Coupe de France                 | beIN Sports + France TV| ✅ diffusable        |
| Équipe de France                | TF1 / M6 (en clair)    | ✅ diffusable        |
| Premier Padel                   | Canal+ (Canal+ Sport)  | ✅ diffusable        |
| Top 14 / rugby                  | Canal+ (Canal+ Sport)  | ✅ diffusable        |
| Formule 1                       | Canal+                 | ✅ diffusable        |
| Serie A                         | DAZN                   | ⛔ hors abonnement   |
| **Liga**                        | **DAZN + Disney+**     | ⛔ hors abonnement   |
| **Ligue 1**                     | **Ligue 1+** (unique)  | ⛔ hors abonnement   |

*(Évolutions 2026-27 : la **Ligue 1 est 100 % sur Ligue 1+** ; la **Liga quitte
beIN** pour DAZN + Disney+ (hors abonnement). beIN conserve la **Ligue 2**
(jusqu'en 2027), la Bundesliga et la Coupe de France — diffusables.)*

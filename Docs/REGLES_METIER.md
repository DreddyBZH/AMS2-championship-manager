# REGLES_METIER.md

# AMS2 Championship Manager

## Règles métier

Version : 1.0

---

# 1. Philosophie générale

Le championnat est construit à partir des manches.

Les données AMS2 ne sont jamais relues pour calculer le championnat.

Le championnat exploite uniquement les feuilles :

```
Manche X
```

Chaque manche est autonome.

---

# 2. Cycle de traitement

Lors d'un import AMS2 :

```
Importer AMS2

↓

CalculerClassementReel()

↓

CalculerClassementGestionIA()

↓

CalculerPoints()

↓

AfficherClassement()

↓

MettreAJourClassementChampionnat()
```

L'ordre est obligatoire.

---

# 3. Classements

Chaque manche possède trois notions de classement.

## 3.1 ClassementReel

Résultat brut provenant d'AMS2.

Aucune règle de gestion n'est appliquée.

Exemples :

```
1
2
3
DNF
DNS
DSQ
```

---

## 3.2 ClassementGestionIA

Classement obtenu après application des règles IA.

Selon la configuration :

```
Les IA restent classées

ou

Les IA deviennent transparentes.
```

Cette colonne est technique.

Elle est masquée.

---

## 3.3 Classement (affiché)

Classement visible par l'utilisateur.

Il correspond toujours à :

```
ClassementGestionIA
```

Cette colonne est la seule utilisée dans l'interface.

---

# 4. États AMS2

Le statut est déterminé ainsi :

Si RacePosition est renseigné :

```
ClassementReel = RacePosition
```

Sinon :

```
Etat vide
        → DNS

DNF
Retired
        → DNF

Disqualified
        → DSQ
```

---

# 5. Gestion des IA

Trois paramètres contrôlent les IA.

## Importer IA

OUI

Les IA sont importées.

NON

Les IA sont ignorées.

---

## Classer IA

OUI

Les IA participent au classement.

NON

Les IA deviennent transparentes.

Les humains remontent dans le classement.

---

## Attribuer des points IA

OUI

Les IA reçoivent les points correspondant à leur classement.

NON

Les IA restent classées mais obtiennent :

```
0 point
```

Le classement est conservé.

Les humains conservent leur rang.

Aucun décalage n'est effectué.

---

# 6. Cas de gestion

## Cas 1

Importer IA = OUI

Classer IA = OUI

Points IA = OUI

Exemple :

```
1 IA

2 Humain

3 Humain

4 IA

5 Humain
```

Points :

```
25

18

15

12

10
```

---

## Cas 2

Importer IA = OUI

Classer IA = OUI

Points IA = NON

Classement :

```
1 IA

2 Humain

3 Humain

4 IA

5 Humain
```

Points :

```
IA      0

Humain 18

Humain 15

IA      0

Humain 10
```

Le classement est conservé.

Seuls les points sont supprimés.

---

## Cas 3

Importer IA = OUI

Classer IA = NON

Points IA = NON

Classement réel :

```
1 IA

2 Humain

3 Humain

4 IA

5 Humain
```

Classement Gestion IA :

```
1 Humain

2 Humain

3 Humain
```

Points :

```
25

18

15
```

Les IA sont totalement transparentes.

---

## Cas 4

Importer IA = NON

Les IA n'existent pas dans le championnat.

---

# 7. Attribution des points

Les points sont calculés à partir du :

```
ClassementGestionIA
```

Le barème est lu dans :

```
Tableau des points
```

Les statuts :

```
DNS

DNF

DSQ
```

utilisent également ce tableau.

Ils peuvent recevoir :

* 0 point

ou

des points

ou

des points négatifs

Le code ne contient aucune valeur en dur.

---

# 8. Championnat

Le championnat est calculé uniquement à partir des manches.

Pour chaque pilote :

* cumul des points

* cumul des P1

* cumul des P2

* cumul des P3

* ballast total

* puissance moyenne

* traînée moyenne

* résultat de chaque manche

---

# 9. Classement général

Le classement est ordonné selon :

1. Points

2. Nombre de P1

3. Nombre de P2

4. Nombre de P3

Les critères complémentaires seront ajoutés ultérieurement.

---

# 10. Interface

Les feuilles modèles sont :

```
MODELE_MANCHE

MODELE_CLASSEMENT
```

Toutes les feuilles sont créées à partir de ces modèles.

Aucune structure ne doit être recréée par le code.

---

# 11. Principes de développement

Le code doit respecter :

* une procédure = une responsabilité

* aucune duplication de règle métier

* aucune valeur métier codée en dur

* lisibilité prioritaire

Les fonctions publiques représentent les traitements métier.

Les fonctions privées représentent les traitements internes.

---

# 12. Évolutions prévues

À terme :

* Pole Position

* Meilleur tour

* Statistiques avancées

* Records

* Sanctions

* Classement équipes

devront s'intégrer sans remettre en cause les règles décrites dans ce document.

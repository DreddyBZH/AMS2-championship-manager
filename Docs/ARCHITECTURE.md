# ARCHITECTURE.md

# AMS2 Championship Manager

## Architecture technique

Version : 1.0

---

# Philosophie

Le projet est développé en VBA sous Microsoft Excel.

L'objectif est de conserver une architecture :

* simple
* modulaire
* lisible
* évolutive

Chaque module possède une responsabilité clairement identifiée.

Les dépendances entre modules doivent rester limitées.

---

# Organisation générale

Le projet est composé de quatre familles.

```
Interface

↓

Import / Export

↓

Moteur métier

↓

Outils
```

---

# Modules Interface

## mdlAccueil

Responsable de :

* création de la page d'accueil
* boutons
* tuiles
* paramètres
* compteurs

Ne contient aucune règle métier.

Ne réalise aucun calcul de championnat.

---

# Modules métier

## mdlChampionship

Responsable du championnat.

Fonctions principales :

* génération des manches
* calcul des classements
* classement général
* classement équipes

Ce module est le cœur de l'application.

---

## mdlStats

Responsable des statistiques.

Exemples :

* victoires
* podiums
* pole positions
* meilleurs tours
* abandon
* kilomètres

Il ne modifie jamais les résultats.

Il exploite uniquement les données existantes.

---

# Import

## mdlAMS2Import

Responsable :

* lecture du JSON AMS2
* import des sessions
* création des pilotes IA
* alimentation des feuilles Manche

Il ne calcule jamais le championnat.

Il ne réalise que les traitements d'import.

---

# Export

## mdlAMS2Export

Responsable :

* génération des exports
* publication
* extraction des données

Aucune règle métier.

---

# Utilitaires

## mdlUtils

Fonctions génériques.

Exemples :

```
GetColumnByHeader()

ToDouble()

SetValeur()

SheetExists()

TrouverLignePilote()

NumeroManche()
```

Aucune fonction métier.

---

# Modèles Excel

Deux modèles sont utilisés.

```
MODELE_MANCHE

MODELE_CLASSEMENT
```

Toutes les feuilles sont créées à partir de ces modèles.

Le code ne doit jamais recréer leur structure.

---

# Flux général

Création du championnat

↓

Création des feuilles Manche

↓

Création du Classement

↓

Import AMS2

↓

Calcul des manches

↓

Calcul du championnat

↓

Calcul des statistiques

↓

Export

---

# Flux d'une manche

```
ImporterSessionQualif()

↓

ImporterSessionCourse()

↓

CalculerClassementReel()

↓

CalculerClassementGestionIA()

↓

CalculerPoints()

↓

AfficherClassement()
```

Cet ordre est obligatoire.

---

# Flux du championnat

```
Lecture des feuilles Manche

↓

Calcul des pilotes

↓

Calcul des équipes

↓

Tri

↓

Affichage
```

Le championnat ne relit jamais les données AMS2.

---

# Dépendances autorisées

```
mdlAccueil
        ↓
mdlChampionship

mdlChampionship
        ↓
mdlUtils

mdlAMS2Import
        ↓
mdlUtils

mdlStats
        ↓
mdlChampionship
```

Les dépendances circulaires sont interdites.

---

# Règles de développement

Une fonction réalise une seule tâche.

Une procédure publique correspond à un traitement métier.

Une procédure privée correspond à un traitement interne.

Une règle métier ne doit exister qu'à un seul endroit.

Les fonctions utilitaires restent totalement génériques.

---

# Convention de nommage

Modules :

```
mdlAccueil

mdlChampionship

mdlAMS2Import
```

Procédures :

```
CalculerPoints()

AfficherClassement()

MettreAJourClassementChampionnat()
```

Variables :

Toujours en français.

Exemples :

```
nbPilotes

ballastTotal

classementReel

classementGestionIA
```

---

# Gestion des erreurs

Les erreurs doivent être traitées au plus près de leur origine.

Éviter les :

```
On Error Resume Next
```

sauf lorsqu'ils sont réellement justifiés.

---

# Interface

La logique métier ne doit jamais dépendre de l'interface.

Inversement, l'interface ne doit contenir aucune règle métier.

---

# Objectif

Le projet doit pouvoir évoluer par ajout de nouvelles fonctionnalités sans remettre en cause les modules existants.

Les évolutions doivent privilégier :

* l'extension

plutôt que

* la réécriture.

Chaque nouvelle fonctionnalité doit naturellement trouver sa place dans cette architecture.

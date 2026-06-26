# AMS2 Championship Manager

## Présentation

AMS2 Championship Manager est un gestionnaire de championnat pour Automobilista 2 développé en VBA sous Microsoft Excel.

L'application permet de :

- Générer automatiquement un championnat
- Importer les résultats AMS2
- Gérer les pilotes humains et les IA
- Calculer les classements
- Appliquer un barème de points configurable
- Générer les statistiques
- Exporter les résultats

Le projet est conçu pour être entièrement configurable sans modifier le code métier.

---

# Architecture

Le projet est organisé en modules VBA.

## Modules principaux

- mdlAccueil
- mdlChampionship
- mdlAMS2Import
- mdlAMS2Export
- mdlStats
- mdlUtils

---

# Structure du classeur

Accueil

Pilote

Equipe

Manches

Manche x

Classement

Classement équipes

Statistiques

Records

Sanctions

---

# Philosophie

Le projet privilégie :

- la lisibilité
- la simplicité
- des fonctions courtes
- une architecture modulaire
- le minimum de duplication

Les règles métier doivent être clairement séparées des traitements d'interface.

---

# Gestion des IA

Trois paramètres pilotent le comportement des IA :

Importer IA

Classer IA

Attribuer des points IA

Le calcul est réalisé selon les règles décrites dans la documentation métier.

---

# Principe général

Import AMS2

↓

Calcul Classement réel

↓

Calcul Classement Gestion IA

↓

Calcul Points

↓

Classement affiché

↓

Classement Championnat

---

# Convention

Toutes les procédures sont en français.

Toutes les variables sont en français.

Option Explicit est obligatoire.

Les procédures publiques représentent les traitements métier.

Les procédures privées représentent les traitements internes.

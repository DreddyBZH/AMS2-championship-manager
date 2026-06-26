# Convention de développement

## Règle n°1

On modifie le minimum de code possible.

On évite les réécritures complètes lorsqu'un correctif suffit.

---

## Règle n°2

Toujours privilégier :

Lisibilité

avant

Optimisation.

---

## Règle n°3

Une procédure = une responsabilité.

Exemple :

CalculerPoints()

ne calcule pas les classements.

CalculerClassementGestionIA()

ne calcule pas les points.

---

## Règle n°4

Les règles métier doivent être centralisées.

Exemple :

Gestion IA

Barème

Statuts DNS/DNF/DSQ

ne doivent jamais être dupliqués.

---

## Règle n°5

Les fonctions utilitaires doivent rester génériques.

Exemple :

GetColumnByHeader()

ToDouble()

SetValeur()

TrouverLignePilote()

ne doivent contenir aucune règle métier.

---

## Règle n°6

Pas de coordonnées en dur.

Toutes les positions de l'interface doivent provenir du module UI.

---

## Règle n°7

Les modèles Excel sont la référence.

MODELE_MANCHE

MODELE_CLASSEMENT

Toute nouvelle feuille est créée à partir de ces modèles.

---

## Règle n°8

Le calcul d'une manche suit obligatoirement l'ordre suivant :

CalculClassementReel()

↓

CalculClassementGestionIA()

↓

CalculPoints()

↓

AfficherClassement()

---

## Règle n°9

Le classement championnat ne relit jamais les données AMS2.

Il exploite uniquement les feuilles "Manche x".

---

## Règle n°10

Les réponses ChatGPT doivent privilégier :

- une fonction complète
- un module complet

plutôt que des snippets.

---

# Méthode de travail

Le développement se fait module par module.

Pour chaque demande :

1. Lecture complète du module.

2. Analyse.

3. Proposition si nécessaire.

4. Validation.

5. Retour du module complet corrigé.

---

# Règle importante

Ne jamais proposer une refonte d'architecture sans demande explicite.

Le projet privilégie les évolutions incrémentales.

---

# Priorités

1. Fiabilité

2. Lisibilité

3. Simplicité

4. Performances

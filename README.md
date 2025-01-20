# Optimisation de Portefeuilles dans des Environnements de Marché

## Description du Projet

Ce projet avait pour objectif de développer un outil permettant de construire et d’évaluer des portefeuilles optimisés selon deux approches principales :  
1. **Portefeuille inconditionnel** : basé sur les rendements moyens et covariances calculés sur l’ensemble de l’historique.  
2. **Portefeuille conditionnel** : basé sur des hypothèses spécifiques à un régime de marché (hausse ou baisse).  

L’outil doit fournir une évaluation précise des performances prévues et effectives des portefeuilles pour des horizons d’investissement à 3 ans, en utilisant des données historiques de rendement.

---

## Fonctionnalités Principales

### Partie 1 : Portefeuilles Inconditionnels
- Construction de portefeuilles basés sur des **rendements moyens** et des **covariances** calculés sur l’ensemble de l’historique.
- Comparaison des portefeuilles obtenus pour trois environnements d’investissement (S&P 500, Stoxx 600, MSCI World).
- Évaluation de la performance effective et analyse des écarts entre prévisions et résultats réels.

### Partie 2 : Portefeuilles Conditionnels
- Prise en compte des **régimes de marché spécifiques** :
  - Régimes de hausse (secteurs offensifs privilégiés).
  - Régimes de baisse (secteurs défensifs privilégiés).
- Analyse des gains en termes de performance et de réduction du risque en remplaçant les portefeuilles inconditionnels par des portefeuilles conditionnels.
- Comparaison des résultats pour des scénarios distincts.

### Partie 3 : Backtest et Évaluation des Performances
- Backtest des portefeuilles sur un historique de 30 ans (360 mois) :
  - Les rendements moyens et les covariances sont calculés sur une période de 72 mois (6 ans) avant chaque point d’investissement.
  - Les performances effectives sont calculées sur une période de 36 mois (3 ans) suivant chaque point d’investissement.
- Maximisation de l’**équivalent certain**.
- Comparaison des performances prévues et effectives, et évaluation des écarts.
- Benchmark : comparaison des portefeuilles optimisés avec un **portefeuille équipondéré** et les **indices de marché**.

---

## Architecture du Projet

Le projet s’appuie sur les concepts suivants :

1. **Feuille de calcul "optimisation" :**  
   - Utilisée pour calculer les allocations optimales en fonction des rendements et covariances prévus.

2. **Feuille de calcul "évaluation" :**  
   - Utilisée pour mesurer les performances effectives des portefeuilles sur les 36 mois suivants.

3. **Analyse comparative :**
   - Identification des écarts entre performances prévues et performances effectives.
   - Étude des implications de ces écarts pour les décisions d’investissement.

4. **Backtest sur un échantillon de 30 ans :**  
   - Les portefeuilles sont recalculés tous les 6 mois pour réduire la charge de calcul.
   - Cela permet d’obtenir environ 42 investissements entre les mois 72 et 324.

---

## Contraintes et Hypothèses

1. **Données historiques :**
   - Les rendements mensuels sont disponibles pour 30 ans (360 mois).
   - Les périodes d’investissement vont de \(t=72\) à \(t=324\).

2. **Période d’optimisation :**
   - Rendements moyens et covariances calculés sur 72 mois (t-71 à t).
   - Prévisions utilisées pour les 36 mois suivants (t+1 à t+36).

3. **Benchmarks :**
   - Comparaison avec un portefeuille équipondéré et les indices de marché (S&P 500, Stoxx 600, MSCI World).

4. **Scénarios de marché :**
   - Régimes de hausse : focus sur secteurs offensifs.
   - Régimes de baisse : focus sur secteurs défensifs.

-

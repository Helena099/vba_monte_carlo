# Guide d'installation du Simulateur Monte-Carlo pour l'Analyse des Risques Financiers

## Prérequis
- Microsoft Excel (version 2010 ou ultérieure)
- Accès à l'éditeur VBA (Alt+F11)
- Macros activées dans Excel

## Étapes d'installation

### 1. Création d'un nouveau classeur Excel
- Ouvrez Excel et créez un nouveau classeur
- Enregistrez-le sous un nom approprié (par exemple "SimulateurMonteCarlo.xlsm")
- Assurez-vous de l'enregistrer au format .xlsm (classeur Excel avec macros)

### 2. Configuration des références VBA
- Appuyez sur Alt+F11 pour ouvrir l'éditeur VBA
- Dans le menu "Outils", sélectionnez "Références"
- Assurez-vous que les références suivantes sont cochées:
  - Visual Basic for Applications
  - Microsoft Excel Object Library
  - Microsoft Forms 2.0 Object Library
  - Microsoft Office Object Library
  - Microsoft Scripting Runtime

### 3. Création des modules VBA
1. **Module principal (mdlMain)**
   - Dans le projet VBA, cliquez droit sur "Modules" -> "Insérer" -> "Module"
   - Renommez le module en "mdlMain"
   - Copiez-collez le code du module mdlMain fourni

2. **Module Monte-Carlo (mdlMonteCarlo)**
   - Créez un nouveau module et nommez-le "mdlMonteCarlo"
   - Copiez-collez le code du module mdlMonteCarlo fourni

3. **Module Graphiques (mdlCharts)**
   - Créez un nouveau module et nommez-le "mdlCharts"
   - Copiez-collez le code du module mdlCharts fourni

4. **Module Rapport (mdlReport)**
   - Créez un nouveau module et nommez-le "mdlReport"
   - Copiez-collez le code du module mdlReport fourni

### 4. Création du formulaire utilisateur
- Dans le projet VBA, cliquez droit sur "UserForms" -> "Insérer" -> "UserForm"
- Renommez le UserForm en "frmMonteCarloSimulator"
- Vous pouvez soit:
  1. **Utiliser le code de création automatique**:
     - Créez un module temporaire
     - Copiez-collez le code de création du formulaire
     - Exécutez la procédure `CreateUserForm()`
     - Supprimez le module temporaire après utilisation
  
  2. **Créer manuellement l'interface**:
     - Ajoutez manuellement tous les contrôles selon le design fourni
     - Configurez les propriétés de chaque contrôle

- Dans le code du formulaire (cliquez sur "Afficher le code"), copiez-collez le code du formulaire fourni

### 5. Configuration d'un point d'entrée
- Créez un bouton sur une feuille Excel ou un élément de menu
- Associez-lui la procédure `InitiateMonteCarloSimulator()` du module mdlMain

## Utilisation du simulateur
1. Cliquez sur le bouton ou l'élément de menu pour lancer le simulateur
2. Entrez les paramètres:
   - Nombre d'itérations (ex: 10 000)
   - Rendement espéré (%)
   - Volatilité (%)
   - Montant initial (€)
   - Horizon temporel (années)
3. Cliquez sur "Lancer la Simulation"
4. Consultez les résultats et les graphiques
5. Exportez un rapport complet si nécessaire

## Dépannage
- Si des erreurs surviennent lors de l'exécution, vérifiez que toutes les références VBA sont correctement configurées
- Assurez-vous que les macros sont activées dans Excel
- Si les graphiques ne s'affichent pas, vérifiez les permissions d'Excel pour la création de graphiques

## Personnalisation
Le code est organisé de manière modulaire, vous permettant de:
- Modifier les algorithmes de simulation dans le module mdlMonteCarlo
- Personnaliser les graphiques dans le module mdlCharts
- Adapter le format du rapport dans le module mdlReport
- Améliorer l'interface utilisateur en modifiant le formulaire
- Ajouter de nouvelles fonctionnalités ou analyses statistiques

## Fonctionnalités avancées
Si vous souhaitez étendre les fonctionnalités du simulateur, voici quelques idées :

1. **Distributions alternatives** : Ajoutez d'autres distributions statistiques pour la génération des rendements (log-normale, Student, etc.)
2. **Modèles stochastiques avancés** : Implémentez des modèles comme le mouvement brownien géométrique ou le modèle de Heston
3. **Backtesting** : Ajoutez une fonctionnalité pour comparer les simulations avec des données historiques
4. **Stress testing** : Permettez à l'utilisateur de définir des scénarios de stress pour tester la robustesse du portefeuille
5. **Optimisation de portefeuille** : Intégrez des méthodes d'optimisation de portefeuille basées sur les résultats des simulations

## Notes techniques
- La transformation Box-Muller est utilisée pour générer des nombres aléatoires suivant une distribution normale
- Le tri à bulles est utilisé pour le calcul de la VaR (Value at Risk) ; pour des ensembles de données très importants, vous pourriez envisager une méthode de tri plus efficace
- Les graphiques sont générés dans des feuilles Excel distinctes puis copiés dans le formulaire VBA
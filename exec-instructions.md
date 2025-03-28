# Instructions pour mettre en place et exécuter le Simulateur Monte-Carlo

## Étape 1 : Préparer le classeur Excel

1. Ouvrez Microsoft Excel
2. Créez un nouveau classeur
3. Appuyez sur `Alt + F11` pour ouvrir l'éditeur VBA
4. Vérifiez que les macros sont activées :
   - Si vous recevez un message d'avertissement concernant les macros, cliquez sur "Activer le contenu"
   - Alternativement, allez dans `Fichier > Options > Centre de gestion de la confidentialité > Paramètres du Centre de gestion de la confidentialité > Paramètres des macros` et sélectionnez "Activer toutes les macros"

## Étape 2 : Importer le code source

### Création des modules
1. Dans l'éditeur VBA, cliquez droit sur le projet dans l'explorateur de projets
2. Sélectionnez `Insérer > Module`
3. Créez les modules suivants :
   - `mdlMain`
   - `mdlMonteCarlo`
   - `mdlCharts`
   - `mdlReport`
4. Copiez le code fourni dans chaque module correspondant

### Création du formulaire utilisateur
1. Dans l'éditeur VBA, cliquez droit sur le projet
2. Sélectionnez `Insérer > UserForm`
3. Renommez-le en `frmMonteCarloSimulator`
4. Vous avez deux options pour créer l'interface :
   - **Option 1** : Création automatique avec le script fourni
     - Créez un module temporaire
     - Copiez-y le code de création du formulaire
     - Exécutez la procédure `CreateUserForm()`
     - Supprimez le module temporaire après utilisation
   - **Option 2** : Création manuelle
     - Ajoutez les contrôles comme indiqué dans la maquette
     - Configurez les propriétés selon les spécifications
5. Une fois le formulaire créé, ajoutez le code du formulaire en faisant un clic droit sur le formulaire et en sélectionnant "Afficher le code"

## Étape 3 : Configurer le point d'entrée

### Option 1 : Bouton sur une feuille Excel
1. Revenez à Excel (Alt+F11 pour sortir de l'éditeur VBA)
2. Allez dans l'onglet `Développeur` (si non visible, activez-le via `Fichier > Options > Personnaliser le ruban`)
3. Cliquez sur `Insérer > Bouton de formulaire`
4. Dessinez le bouton sur la feuille
5. Dans la boîte de dialogue "Assigner une macro", sélectionnez `InitiateMonteCarloSimulator`
6. Renommez le bouton en "Lancer le Simulateur Monte-Carlo"

### Option 2 : Raccourci clavier
1. Dans l'éditeur VBA, créez une nouvelle procédure dans `mdlMain` :
```vba
Sub LauncherShortcut()
    Call InitiateMonteCarloSimulator
End Sub
```
2. Assignez un raccourci clavier : `Outils > Options > Éditeur > Raccourcis clavier`

## Étape 4 : Test et exécution

1. Enregistrez le classeur au format `.xlsm` (classeur Excel avec macros)
2. Fermez l'éditeur VBA
3. Cliquez sur le bouton créé ou utilisez le raccourci clavier
4. Le formulaire du simulateur devrait apparaître
5. Entrez les paramètres souhaités :
   - Nombre d'itérations (ex: 10 000)
   - Rendement espéré (%)
   - Volatilité (%)
   - Montant initial (€)
   - Horizon temporel (années)
6. Cliquez sur "Lancer la Simulation"
7. Observez les résultats et les visualisations générées

## Conseils pour l'utilisation

- Pour des simulations rapides, commencez avec un nombre d'itérations moins élevé (ex: 1 000)
- La VaR (Value at Risk) à 95% indique le montant minimal auquel votre portefeuille pourrait tomber avec une confiance de 95%
- Utilisez les différents onglets de visualisation pour mieux comprendre la distribution des résultats
- Le rapport généré peut être imprimé ou sauvegardé en PDF

## Résolution des problèmes courants

- **Erreur "Macro non disponible"** : Vérifiez que le code a été correctement importé dans tous les modules
- **Graphiques manquants** : Assurez-vous que les feuilles de calcul requises ont été créées
- **Performance lente** : Réduisez le nombre d'itérations ou optimisez le code de tri
- **Résultats incohérents** : Vérifiez les paramètres d'entrée, en particulier la volatilité et le rendement attendu

## Notes importantes

- Les performances passées ne garantissent pas les résultats futurs
- Ce simulateur est un outil d'aide à la décision et ne constitue pas un conseil financier
- Pour des décisions d'investissement importantes, consultez un conseiller financier professionnel
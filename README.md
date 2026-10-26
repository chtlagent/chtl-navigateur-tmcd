# CHTL — Navigateur TMCD

Complément local pour Excel Windows développé en C# / VSTO (.NET Framework 4.8), avec un véritable volet de tâches à droite d'Excel.

## Confidentialité

**Ne jamais déposer sur GitHub de classeurs hospitaliers, fichiers de dotations, exports BCA, données patient, accès réseau ou informations d'identification.** Le projet ne contient que du code d'interface, des noms de feuilles et la logique de navigation. Le volet fonctionne localement et n'effectue pas d'appels réseau.

## Fonctionnalités

- Volet ancré à droite, visible lorsqu'un classeur possède les feuilles `DOT_Base` et `DOT - Contrôle`.
- Liste dynamique des feuilles du classeur (y compris celles ajoutées après compilation).
- Recherche et classement en familles ; feuilles masquées affichables dans la liste mais **jamais démasquées** automatiquement.
- Boutons Accueil, Précédent, Contrôle DOT et Effacer filtres.
- Effacer filtres appelle la macro existante `Reinitialiser_Filtres_DOT` dans le classeur ouvert.
- Ruban Excel : Compléments > CHTL | Pilotage > Navigateur TMCD.

## Compiler sans Visual Studio sur le PC CHTL

1. Ouvrir l'onglet **Actions** du dépôt privé.
2. Sélectionner **Compiler le Navigateur TMCD VSTO**, puis éventuellement **Run workflow**.
3. Le workflow utilise `windows-2022` et MSBuild de Visual Studio 2022, et cherche les cibles de compilation VSTO.
4. Si le Build réussit, récupérer l'artefact `CHTL-Navigateur-TMCD-VSTO-NON-SIGNE`.
5. **Attention : cet artefact est un diagnostic de compilation, pas un installateur prêt à déployer.**

## Installation sur le poste du CHTL

Une fois la compilation confirmée, créer un paquet ClickOnce **signé par un certificat accepté par l'organisation**, valider la compatibilité Office/architecture et faire approuver le déploiement par le service informatique. Le runtime VSTO est déjà détecté sur le poste ; cela n'annule pas la nécessité de signer et déployer correctement l'add-in.

## Validation fonctionnelle

Tester sur un classeur factice (sans données CHTL) : création du volet, recherche, navigation, retour, 50+ feuilles et appel à une macro test. Le code n'a pas encore été testé en exécution réelle dans Excel.

## Dossier principal

`src/CHTL_Navigateur_TMCD/CHTL_Navigateur_TMCD.csproj` — projet Visual Studio compatible avec MSBuild 2022.

## Note technique

Les fichiers `ThisAddIn.Designer.cs` et `ThisAddIn.Designer.xml` proviennent d'un modèle VSTO adapté ; ils doivent être vérifiés par la compilation et peuvent nécessiter une régénération dans Visual Studio 2022 si Microsoft modifie le modèle.

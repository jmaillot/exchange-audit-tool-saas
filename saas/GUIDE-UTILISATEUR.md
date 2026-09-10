# Guide utilisateur — M365 Audit Tool

L'interface est disponible en français et en anglais : bouton **FR/EN** en haut à droite (le choix est mémorisé). Les noms de propriétés Exchange restent toujours en anglais car ce sont les noms techniques (colonnes PowerShell/CSV).

## 1. Première connexion

1. Ouvrez l'application et allez sur **Connection**.
2. Tapez votre **adresse e-mail professionnelle** (UPN, ex. `vous@contoso.onmicrosoft.com`).
3. **Première fois dans votre organisation ?** Cliquez le lien **Register this tenant** dans le bandeau bleu (un administrateur doit valider, une seule fois par tenant), puis revenez ici.
4. Cliquez **Se connecter avec Microsoft** et connectez-vous avec un compte lecteur Exchange (ex. rôle *View-Only Organization Management*).
5. Le bandeau en haut passe au vert. Le jeton d'accès reste **dans l'onglet uniquement** (survit à F5, jamais sur disque) : **Disconnect** ou la fermeture de l'onglet le jette.

## 2. Lancer un audit

1. Dans le menu de gauche, ouvrez une catégorie puis une section (ex. *Mailboxes → User mailboxes*).
2. Cochez les propriétés voulues (la case **Filtrer** aide à les retrouver, **Tout sélectionner** tout coche — les noms de propriétés restent en anglais car ce sont les noms Exchange). Les options marquées **lent** interrogent chaque objet une par une : c'est beaucoup plus long, décochez-les pour un premier passage.
3. Le bloc **Mode intelligent** (coché par défaut, recommandé) ne garde que les colonnes remplies sur au moins une ligne.
4. Cliquez **LANCER L'AUDIT**. La progression s'affiche dans *Aperçu des résultats* (200 premières lignes, avec chrono).
5. **Télécharger CSV / Télécharger XLSX** pour récupérer le fichier (`NomSection-date.csv`). Les colonnes supprimées par le mode intelligent sont listées dans le journal sous le tableau.

## 3. Problèmes fréquents

| Symptôme | Solution |
|---|---|
| `Sign-in failed` à la connexion | Le tenant n'est pas enregistré : passez par **Register this tenant** (admin), puis reconnectez-vous. |
| Job `failed - worker 500: UnAuthorized` | Consentement admin manquant **ou** compte sans rôle lecteur Exchange. Refaites le consentement, **Disconnect + Connect** (le jeton doit être réémis), relancez. |
| Cases à cocher invisibles après une mise à jour | **Ctrl+F5** (refresh forcé) une fois. |
| `Session expired` au lancement | Reconnectez-vous (le jeton a expiré). |
| Audit très long | Décochez les options **slow**, ou limitez avec **Result size: First 1000**. |

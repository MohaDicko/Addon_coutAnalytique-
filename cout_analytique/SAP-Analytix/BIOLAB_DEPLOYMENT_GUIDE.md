# Guide de Déploiement sur BIOLAB

## 🔐 Accès BIOLAB

```
RDP Host: biolab-win.safirconsulting.net
Utilisateur: safir
Mot de passe: safir2
```

## 📦 Étape 1 : Préparer le Projet pour BIOLAB

Le projet est prêt dans : `C:\Users\safir\source\repos\SAP-Analytix\`

Fichiers essentiels :
- ✅ `SAP_Analytix/` - Dossier du projet
- ✅ `SAP_Analytix.sln` - Solution Visual Studio
- ✅ `Program.cs` - Code source refactorisé
- ✅ `SAP_Analytix.xml` - Manifest de l'add-on

## 🚀 Étape 2 : Transférer vers BIOLAB

### Option A : Via RDP (Recommandé)
1. Connectez-vous à BIOLAB via RDP
2. Ouvrez l'explorateur fichiers
3. Naviguez à : `C:\Users\safir\source\repos\SAP-Analytix\`
4. Copiez le dossier entier

### Option B : Via Git (Si Git est installé sur BIOLAB)
```powershell
# Sur BIOLAB
cd C:\Users\safir\source\repos\
git clone <votre_repo_url>
```

## 🔨 Étape 3 : Compiler sur BIOLAB

1. **Ouvrez le projet** dans Visual Studio (présent sur BIOLAB)
2. **Ouvrez la solution** : `SAP_Analytix.sln`
3. **Restaurez les packages NuGet** : `Tools → NuGet → Restore Packages`
4. **Compilez** : `Ctrl + Shift + B`

La compilation devrait réussir puisque :
- ✅ SAP B1 10.0 est installé
- ✅ SDK est disponible
- ✅ DLL COM sont enregistrées

## ⚙️ Étape 4 : Enregistrer l'Add-on dans SAP B1

1. **Ouvrez SAP Business One** sur BIOLAB
2. **Allez à** : `Tools → Add-Ons & Macros → Add-Ons`
3. **Cliquez sur** : `Settings`
4. **Enregistrez le nouvel add-on** :
   - Chemin : `C:\Users\safir\source\repos\SAP-Analytix\SAP_Analytix\bin\Debug\SAP_Analytix.exe`
   - Nom : `SAP_Analytix`
   - Version : `1.0.0`
   - Description : `Gestion automatique des centres de coûts`

## 🧪 Étape 5 : Tester sur BIOLAB

### Test 1 : Commande Client
1. Ouvrez `Ventes → Commandes clients`
2. Créez une nouvelle commande
3. Sélectionnez un article dans la matrice
4. **Vérifiez** : Le centre de coûts doit être rempli automatiquement ✅

### Test 2 : Commande Fournisseur
1. Ouvrez `Achats → Commandes fournisseurs`
2. Créez une nouvelle commande
3. Sélectionnez un article
4. **Vérifiez** : Le centre de coûts doit être rempli automatiquement ✅

### Test 3 : Facture Client
1. Ouvrez `Ventes → Factures clients`
2. Créez une nouvelle facture
3. Sélectionnez un article
4. **Vérifiez** : Le centre de coûts doit être rempli automatiquement ✅

### Test 4 : Autres Documents
Répétez pour :
- ☐ Réception de marchandises (GRN)
- ☐ Livraison client
- ☐ Facture fournisseur

## ✅ Critères de Succès

- [x] Compilation sans erreurs
- [ ] L'add-on se lance sans crash
- [ ] Les centres de coûts se remplissent automatiquement
- [ ] Le champ fonctionne sur tous les types de documents
- [ ] Pas d'erreurs dans la barre de statut SAP

## 🔧 Débogage sur BIOLAB

Si erreurs lors de l'exécution :

1. **Vérifiez la barre de statut SAP** : Elle affiche les messages d'erreur
2. **Consultez l'Event Viewer Windows** (`eventvwr.msc`)
3. **Vérifiez le chemin** de la fiche article : `Fichier de stock → Gestion des articles`
4. **Vérifiez l'UDF** : `Administration → Personnalisation → Fiches utilisateur`
   - Cherchez `U_Analytique` dans OITM

## 📋 Checklist Finale

- [ ] Projet compilé sans erreurs
- [ ] Add-on enregistré dans SAP B1
- [ ] Tests passés sur tous les documents
- [ ] Messages de succès affichés
- [ ] Prêt pour production

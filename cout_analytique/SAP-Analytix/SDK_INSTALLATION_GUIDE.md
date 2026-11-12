# Guide d'Installation SDK SAP Business One

## 📋 Prérequis
- Visual Studio Community 2026 (✅ Vous l'avez)
- Accès à une installation SAP Business One 10.0 HANA
- Droits administrateur sur le poste

## 🔧 Étape 1 : Localisez le SDK SAP

Le SDK est installé lors de l'installation de SAP Business One. Cherchez le dossier :

**Chemins typiques :**
- `C:\Program Files (x86)\SAP\SAP Business One SDK\`
- `C:\Program Files\SAP\SAP Business One SDK\`
- `C:\SAP\SDK\`

## ✅ Étape 2 : Enregistrez les DLL COM

Les DLL SAP doivent être enregistrées auprès de Windows. Ouvrez **PowerShell en tant qu'administrateur** et exécutez :

```powershell
# Enregistrer SAPbouiCOM.dll
regsvr32 "C:\Program Files (x86)\SAP\SAP Business One SDK\Lib\SAPbouiCOM.dll"

# Enregistrer SAPbobsCOM.dll
regsvr32 "C:\Program Files (x86)\SAP\SAP Business One SDK\Lib\SAPbobsCOM.dll"
```

## 🔗 Étape 3 : Ajouter les références dans Visual Studio

1. **Ouvrez le projet** `SAP-Analytix.sln` dans Visual Studio 2026
2. **Clic droit** sur le projet `SAP_Analytix` → **Add Reference...**
3. Allez à l'onglet **COM**
4. Recherchez et sélectionnez :
   - ☑️ `SAP Business One UI Object Model` (SAPbouiCOM)
   - ☑️ `SAP Business One DI Object Model` (SAPbobsCOM)
5. Cliquez **Add** puis **OK**

Les références doivent maintenant apparaître dans votre projet.

## ⚠️ Étape 4 (Si références pas trouvées)

Si les références n'apparaissent pas en COM, ajoutez-les manuellement :

1. **Clic droit** sur le projet → **Add Reference...**
2. Allez à **Browse**
3. Naviguez à : `C:\Program Files (x86)\SAP\SAP Business One SDK\Lib\`
4. Sélectionnez :
   - `SAPbouiCOM.dll`
   - `SAPbobsCOM.dll`
5. Cliquez **Add** puis **OK**

## 🔨 Étape 5 : Compilez le projet

```
Ctrl + Shift + B    (ou Build → Build Solution)
```

Si la compilation réussit, tous les erreurs liées aux DLL SAP sont résolues ! ✅

## 🐛 Dépannage

| Erreur | Solution |
|--------|----------|
| `DLL not found` | Vérifiez le chemin d'installation du SDK |
| `COM reference not registered` | Exécutez `regsvr32` en tant qu'administrateur |
| `Type not found` | Assurez-vous que `EmbedInteropTypes=True` est défini |

## 📦 Chemins d'Installation Alternatifs

Si BIOLAB utilise un autre chemin, demandez à Emmanuel Dosso :

```
Chemin SDK sur BIOLAB : _________________________
```

Accès RDP : `biolab-win.safirconsulting.net`

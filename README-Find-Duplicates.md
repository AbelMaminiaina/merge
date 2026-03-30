# Find Duplicates - Scripts PowerShell

Scripts pour trouver les doublons dans les colonnes Excel et exporter toutes les lignes correspondantes.

## Prérequis

- Windows avec PowerShell
- Microsoft Excel installé

## Scripts disponibles

- **Find-Duplicates-NomFR.ps1** : Analyse la colonne "Nom FR"
- **Find-Duplicates-Trigramme.ps1** : Analyse la colonne "trigramme"
- **find-duplicates-all.bat** : Exécute les deux analyses

## Utilisation

### Méthode 1 : Script batch (recommandé)

```batch
find-duplicates-all.bat "C:\chemin\vers\fichier.xlsx"
```

Ou glisser-déposer le fichier Excel sur `find-duplicates-all.bat`

### Méthode 2 : Scripts PowerShell individuels

```powershell
# Analyser Nom FR
.\Find-Duplicates-NomFR.ps1 -FilePath "C:\chemin\vers\fichier.xlsx"

# Analyser trigramme
.\Find-Duplicates-Trigramme.ps1 -FilePath "C:\chemin\vers\fichier.xlsx"
```

## Paramètres

| Paramètre | Description | Défaut |
|-----------|-------------|--------|
| `-FilePath` | Chemin du fichier Excel (obligatoire) | - |
| `-SheetName` | Nom de l'onglet | "Objets de gestion" |
| `-ColumnName` | Nom de la colonne à analyser | "Nom FR" ou "trigramme" |
| `-OutputPath` | Fichier de sortie | `[fichier]_Doublons_[colonne].xlsx` |

## Exemple de personnalisation

```powershell
.\Find-Duplicates-NomFR.ps1 `
    -FilePath "fichier.xlsx" `
    -SheetName "Tables" `
    -ColumnName "Code" `
    -OutputPath "doublons_code.xlsx"
```

## Résultats

Le script crée un fichier Excel contenant :
- **En-tête** : En gras avec fond bleu clair
- **Données** : Toutes les lignes contenant des doublons, triées par valeur
- **Filtres** : Activés automatiquement
- **Colonnes** : Largeur ajustée automatiquement

### Exemple pour IBIA_ACCMGR_Outil.xlsx

**Nom FR :** 8 doublons, 16 lignes exportées
- Date, edition, garanti, générique, option, réel, refus, solidaire

**trigramme :** 87 doublons, 187 lignes exportées
- ABR, ACG, AUT, GRP, MGR, PAY, etc.

## Note importante

Les dates peuvent apparaître comme des nombres dans le fichier de sortie. Pour corriger :
1. Ouvrir le fichier généré
2. Sélectionner la colonne de date
3. Clic droit > Format de cellule > Date
4. Choisir le format souhaité

## Gestion des erreurs

Le script affiche des messages clairs en cas de :
- Fichier inexistant
- Onglet non trouvé (liste les onglets disponibles)
- Colonne non trouvée
- Aucun doublon trouvé

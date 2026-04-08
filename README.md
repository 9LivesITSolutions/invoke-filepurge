# Invoke-FilePurge.ps1

Script PowerShell de **purge automatisée de fichiers anciens**, conçu pour la production et les tâches planifiées Windows. Compatible PS 5.1 et PS 7+.

---

## Fonctionnalités

- Filtrage par âge (`LastWriteTime` ou `CreationTime`), extensions, expressions régulières
- **Mode simulation** (`-WhatIf`) sans aucune suppression réelle
- Log structuré horodaté avec **rotation automatique**
- Rapport CSV des fichiers supprimés
- **Coupe-circuit** : quota volume et quota fichiers par exécution
- Suppression optionnelle des **dossiers vides** après purge
- Écriture dans le **journal d'événements Windows** (Application)
- Codes de sortie normalisés pour le **monitoring** de la tâche planifiée
- Moteur d'énumération .NET natif — testé sur **8 M+ fichiers**

---

## Prérequis

| Élément | Version minimale |
|---|---|
| PowerShell | 5.1 (Windows) ou 7+ |
| OS | Windows Server 2016+ / Windows 10+ |
| Droits | Lecture sur le chemin cible, Écriture pour la suppression |
| Journal Windows | Droits admin requis pour créer une source d'événement |

---

## Installation

```powershell
# Copier le script dans le dossier de scripts
Copy-Item Invoke-FilePurge.ps1 C:\Scripts\

# Débloquer si téléchargé depuis le réseau
Unblock-File -Path C:\Scripts\Invoke-FilePurge.ps1
```

---

## Paramètres

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-TargetPath` | `string[]` | **Obligatoire** | Un ou plusieurs chemins racine à purger |
| `-AgeDays` | `int` | `90` | Âge minimum des fichiers en jours |
| `-UseCreationTime` | `switch` | — | Utilise `CreationTime` au lieu de `LastWriteTime` |
| `-IncludeExtensions` | `string[]` | `@()` (toutes) | Extensions à inclure (ex. `.log`, `.tmp`) |
| `-ExcludeExtensions` | `string[]` | `@()` (aucune) | Extensions à exclure explicitement |
| `-ExcludePatterns` | `string[]` | `@()` (aucun) | Regex appliquées sur le chemin complet pour exclure |
| `-MaxDeleteMB` | `long` | `10240` | Volume max supprimé par exécution en Mo (10 Go) |
| `-MaxFiles` | `long` | `500000` | Nombre max de fichiers supprimés par exécution |
| `-LogPath` | `string` | Dossier du script | Dossier de destination des fichiers log |
| `-LogRetentionDays` | `int` | `30` | Durée de conservation des logs en jours |
| `-PurgeEmptyFolders` | `switch` | — | Supprime les dossiers vides après purge |
| `-WhatIf` | `switch` | — | Simulation : liste sans supprimer |
| `-WriteEventLog` | `switch` | — | Écrit un événement dans le journal Windows |
| `-EventSource` | `string` | `FilePurge` | Nom de la source dans le journal Windows |

---

## Exemples

### 1. Simulation — vérifier avant de supprimer

Toujours commencer par un `-WhatIf` pour valider les candidats sans toucher aux fichiers.

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Logs" `
    -AgeDays 90 `
    -WhatIf
```

Le log `[DEBUG] Diagnostic LastWriteTime` affiche les dates min/max rencontrées et le seuil appliqué.

---

### 2. Purge simple — logs IIS de plus de 90 jours

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "C:\inetpub\logs\LogFiles" `
    -AgeDays 90 `
    -IncludeExtensions '.log' `
    -LogPath "C:\Admin\Logs"
```

---

### 3. Purge multi-chemins avec exclusions

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Logs", "E:\Temp", "F:\Archives\Import" `
    -AgeDays 60 `
    -IncludeExtensions '.log', '.tmp', '.bak' `
    -ExcludePatterns 'KEEP_', '_PERMANENT', '\\audit\\' `
    -LogPath "D:\Admin\Purge\Logs" `
    -WriteEventLog
```

> Les patterns `-ExcludePatterns` sont des **expressions régulières** appliquées sur le chemin complet.  
> Exemples : `'\\audit\\'` exclut tout fichier sous un dossier `audit`, `'KEEP_'` exclut les fichiers dont le chemin contient `KEEP_`.

---

### 4. Purge avec quota augmenté et dossiers vides

Cas typique : rattrapage initial sur un volume avec des années d'accumulation.

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Archives" `
    -AgeDays 365 `
    -MaxDeleteMB 51200 `
    -MaxFiles 1000000 `
    -PurgeEmptyFolders `
    -LogPath "C:\Admin\Logs" `
    -WriteEventLog
```

---

### 5. Basé sur CreationTime au lieu de LastWriteTime

Utile quand les fichiers sont copiés régulièrement (LastWriteTime réinitialisé) mais la date de création reste fiable.

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "E:\Exports\Daily" `
    -AgeDays 30 `
    -UseCreationTime `
    -IncludeExtensions '.csv', '.xlsx' `
    -LogPath "C:\Admin\Logs"
```

---

### 6. Tâche planifiée Windows — configuration recommandée

**Action de la tâche planifiée :**

| Champ | Valeur |
|---|---|
| Programme | `powershell.exe` |
| Arguments | voir ci-dessous |
| Démarrer dans | `C:\Scripts` |

```
powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -TargetPath "D:\Logs" -AgeDays 90 -MaxDeleteMB 20480 -LogPath "C:\Admin\Logs" -WriteEventLog
```

**Configurer une alerte sur le code de sortie** dans votre outil de monitoring :

```powershell
# Exemple : vérifier le dernier code de sortie depuis un script de supervision
$result = Start-Process powershell.exe -ArgumentList '...' -Wait -PassThru
switch ($result.ExitCode) {
    0 { Write-Host "OK" }
    1 { Send-Alert "CRITIQUE : erreur de purge" }
    2 { Send-Alert "AVERTISSEMENT : quota atteint, purge partielle" }
    3 { Send-Alert "AVERTISSEMENT : erreurs sur certains fichiers" }
}
```

---

## Codes de sortie

| Code | Signification | Action recommandée |
|---|---|---|
| `0` | Succès complet | — |
| `1` | Erreur critique (chemin invalide, permissions) | Vérifier le log, corriger les droits |
| `2` | Quota atteint, purge partielle | Augmenter `-MaxDeleteMB` ou planifier plus fréquemment |
| `3` | Avertissement : erreurs sur certains fichiers | Consulter le log pour les fichiers en erreur |

---

## Fichiers générés

| Fichier | Description |
|---|---|
| `FilePurge_YYYYMMDD_HHMMSS.log` | Log structuré complet de l'exécution |
| `FilePurge_YYYYMMDD_HHMMSS_report.csv` | CSV des fichiers supprimés (chemin, âge, taille, statut) |

### Format du log

```
2026-01-15 03:00:01 === [SECTION] ======================================================================
2026-01-15 03:00:01 === [SECTION] INVOKE-FILEPURGE v2.5.0  —  MODE RÉEL
2026-01-15 03:00:01     [INFO]    Âge minimum : 90 jours (LastWriteTime) — seuil : 2025-10-17
2026-01-15 03:00:01 ... [DEBUG]   Moteur d'énumération : récursion manuelle .NET Framework (PS5.1)
2026-01-15 03:00:45 ... [DEBUG]   Diagnostic LastWriteTime — plus ancien : 2023-04-02 | plus récent : 2026-01-14 | seuil purge : 2025-10-17
2026-01-15 03:01:12 [+] [SUCCESS] Supprimé : D:\Logs\app_20230402.log  (age: 653j, 2,14 MB)
2026-01-15 03:02:00     [INFO]    Fichiers scannés : 45 231 — Candidats : 12 847 — Supprimés : 12 847
```

### Format du rapport CSV

```csv
"Path","AgeDays","SizeBytes","DeletedAt","Status"
"D:\Logs\app_20230402.log","653","2244608","2026-01-15T03:01:12","Deleted"
"D:\Logs\app_20230403.log","652","1887232","2026-01-15T03:01:12","Deleted"
"D:\Temp\import_20230101.tmp","379","512","2026-01-15T03:01:13","Error: Access denied"
```

---

## Comportement du quota

Le quota est un **coupe-circuit de sécurité**, pas un objectif. Quand il est atteint :

- La purge s'arrête immédiatement (les anciens fichiers sont prioritaires — tri par date croissante)
- Le code de sortie passe à `2`
- La ligne `[WARN] Quota volume atteint` apparaît dans le log
- Un événement Windows de type `Warning` est émis si `-WriteEventLog` est activé

Pour un rattrapage initial sur un volume chargé, planifier plusieurs exécutions successives ou augmenter temporairement `-MaxDeleteMB`.

---

## Compatibilité PS 5.1 vs PS 7+

Le script détecte automatiquement le runtime et choisit le moteur d'énumération optimal :

| Runtime | Moteur | Caractéristique |
|---|---|---|
| PS 7+ / .NET 5+ | `EnumerationOptions` | Plus rapide, `IgnoreInaccessible` natif |
| PS 5.1 / .NET Framework 4.x | Récursion `TopDirectoryOnly` | Robuste aux dossiers inaccessibles |

La ligne `[DEBUG] Moteur d'énumération :` dans le log confirme le moteur utilisé.
# Invoke-FilePurge.ps1

> 🇬🇧 [English version available](README.md)

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![Version](https://img.shields.io/badge/Version-2.5.0-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

Script PowerShell de **purge automatisée de fichiers anciens** pour les environnements Windows. Conçu pour les tâches planifiées et les systèmes de fichiers volumineux (testé sur 8 M+ fichiers). Compatible PS 5.1 et PS 7+.

---

## Fonctionnalités

- Filtrage par âge (`LastWriteTime` ou `CreationTime`), extensions, expressions régulières d'exclusion
- **Mode simulation** (`-WhatIf`) — liste les candidats sans rien supprimer
- Log structuré horodaté avec **rotation automatique**
- **Rapport CSV** des fichiers supprimés (chemin, âge, taille, statut)
- **Coupe-circuit** — quota volume et quota fichiers configurables par exécution
- Suppression optionnelle des **dossiers vides** après purge
- Intégration au **journal d'événements Windows** (journal Application)
- **Codes de sortie** normalisés pour le monitoring de la tâche planifiée
- Moteur d'énumération .NET natif — testé sur **8 M+ fichiers**

---

## Prérequis

| Élément | Minimum |
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
| `-ExcludePatterns` | `string[]` | `@()` (aucun) | Regex appliquées sur le chemin complet |
| `-MaxDeleteMB` | `long` | `10240` | Volume max supprimé par exécution en Mo (10 Go) |
| `-MaxFiles` | `long` | `500000` | Nombre max de fichiers supprimés par exécution |
| `-LogPath` | `string` | Dossier du script | Dossier de destination des fichiers log |
| `-LogRetentionDays` | `int` | `30` | Durée de conservation des logs en jours |
| `-PurgeEmptyFolders` | `switch` | — | Supprime les dossiers vides après purge |
| `-WhatIf` | `switch` | — | Mode simulation — aucune suppression |
| `-WriteEventLog` | `switch` | — | Écrit un événement dans le journal Windows |
| `-EventSource` | `string` | `FilePurge` | Nom de la source dans le journal Windows |

---

## Exemples

### 1. Simulation — toujours vérifier avant de supprimer

Lancer `-WhatIf` en premier pour valider les candidats sans toucher aux fichiers.
La ligne `[DEBUG] Diagnostic LastWriteTime` dans le log affiche les dates min/max rencontrées et le seuil appliqué.

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Logs" `
    -AgeDays 90 `
    -WhatIf
```

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

> `-ExcludePatterns` sont des **expressions régulières** appliquées sur le chemin complet.
> Exemples : `'\\audit\\'` exclut tout fichier sous un dossier `audit`, `'KEEP_'` exclut les fichiers dont le chemin contient `KEEP_`.

---

### 4. Rattrapage initial avec quota augmenté et suppression des dossiers vides

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

### 5. Purge basée sur CreationTime

Utile quand les fichiers sont régulièrement copiés (`LastWriteTime` réinitialisé) mais que la date de création reste fiable.

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

**Configuration de l'action de la tâche :**

| Champ | Valeur |
|---|---|
| Programme | `powershell.exe` |
| Arguments | voir ci-dessous |
| Démarrer dans | `C:\Scripts` |

```
powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -TargetPath "D:\Logs" -AgeDays 90 -MaxDeleteMB 20480 -LogPath "C:\Admin\Logs" -WriteEventLog
```

**Surveillance du code de sortie** dans votre outil de monitoring :

```powershell
$result = Start-Process powershell.exe -ArgumentList '...' -Wait -PassThru
switch ($result.ExitCode) {
    0 { Write-Host "OK -- purge terminée" }
    1 { Send-Alert "CRITIQUE : échec de la purge (chemin invalide ou permissions)" }
    2 { Send-Alert "AVERTISSEMENT : quota atteint, purge partielle" }
    3 { Send-Alert "AVERTISSEMENT : erreurs sur certains fichiers, vérifier le log" }
}
```

---

## Codes de sortie

| Code | Signification | Action recommandée |
|---|---|---|
| `0` | Succès complet | — |
| `1` | Erreur critique (chemin invalide, permissions) | Vérifier le log, corriger les droits |
| `2` | Quota atteint, purge partielle | Augmenter `-MaxDeleteMB` ou planifier plus fréquemment |
| `3` | Avertissement — erreurs sur des fichiers individuels | Consulter le log pour le détail des erreurs |

---

## Fichiers générés

| Fichier | Description |
|---|---|
| `FilePurge_YYYYMMDD_HHMMSS.log` | Log structuré complet de l'exécution |
| `FilePurge_YYYYMMDD_HHMMSS_report.csv` | Rapport CSV des fichiers supprimés |

### Format du log

Les messages du log sont en **anglais**.

```
2026-01-15 03:00:01 === [SECTION] ===============================================================
2026-01-15 03:00:01 === [SECTION] INVOKE-FILEPURGE v2.5.0  --  REAL MODE
2026-01-15 03:00:01     [INFO]    Minimum age : 90 days (LastWriteTime) -- cutoff: 2025-10-17
2026-01-15 03:00:01 ... [DEBUG]   Enumeration engine: manual .NET Framework recursion (PS5.1)
2026-01-15 03:00:45 ... [DEBUG]   Diagnostic LastWriteTime -- oldest: 2023-04-02 | newest: 2026-01-14 | cutoff: 2025-10-17
2026-01-15 03:01:12 [+] [SUCCESS] Deleted: D:\Logs\app_20230402.log  (age: 653d, 2.14 MB)
2026-01-15 03:02:00     [INFO]    Scan complete: 45,231 files, 12,847 candidates for purge.
```

**Niveaux de log :**

| Icône | Niveau | Signification |
|---|---|---|
| ` ` | `INFO` | Information générale |
| `...` | `DEBUG` | Détail technique (diagnostic, progression) |
| `[+]` | `SUCCESS` | Fichier ou dossier supprimé avec succès |
| `[!]` | `WARN` | Avertissement (quota, erreur non bloquante) |
| `[X]` | `ERROR` | Erreur bloquante |
| `===` | `SECTION` | Séparateur de section |

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

- La purge s'arrête immédiatement — les fichiers les plus anciens sont traités en priorité (tri par date croissante)
- Le code de sortie passe à `2`
- La ligne `[WARN] Volume quota reached` apparaît dans le log
- Un événement Windows de type `Warning` est émis si `-WriteEventLog` est activé

Pour un rattrapage initial sur un volume chargé, augmenter temporairement `-MaxDeleteMB` ou planifier plusieurs exécutions successives.

---

## Compatibilité PS 5.1 / PS 7+

Le script détecte automatiquement le runtime et choisit le moteur d'énumération optimal :

| Runtime | Moteur | Caractéristique |
|---|---|---|
| PS 7+ / .NET 5+ | `EnumerationOptions` | Plus rapide, `IgnoreInaccessible` natif |
| PS 5.1 / .NET Framework 4.x | Récursion `TopDirectoryOnly` manuelle | Robuste aux dossiers inaccessibles |

La ligne `[DEBUG] Enumeration engine:` dans le log confirme le moteur actif.

---

## Changelog

Voir [CHANGELOG.md](CHANGELOG.md) pour l'historique complet des versions.

---

## Licence

[MIT](LICENSE) — © 2026 9 Lives IT Solutions

# Invoke-FilePurge.ps1

> 🇬🇧 [English version available](README.md)

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![Version](https://img.shields.io/badge/Version-3.3.4-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

Script PowerShell de **purge automatisée de fichiers** pour les environnements Windows. Moteur de règles JSON par chemin, traitement parallèle (PS7+), conçu pour les tâches planifiées. Compatible PS 5.1 et PS 7+.

---

## Fonctionnalités

- **Moteur de règles JSON** — filtres, quotas et options par chemin dans un fichier de configuration unique
- **Mode CLI rétrocompatible** — tous les paramètres v2.x fonctionnent toujours
- **Traitement parallèle** (`-Parallel`) — exécution concurrente des règles sous PS7+, journalisation thread-safe
- Filtrage par âge (`LastWriteTime` ou `CreationTime`), extensions, expressions régulières
  - `IncludeNamePatterns` — regex sur le **nom de fichier**
  - `IncludePathPatterns` — regex sur le **chemin complet**
  - `ExcludePatterns` — regex sur le **chemin complet**
- **Mode simulation** (`-WhatIf`) — liste les candidats sans rien supprimer
- Log structuré horodaté avec **rotation automatique**
- **Rapport CSV** des fichiers supprimés (chemin, âge, taille, statut)
- **Coupe-circuit** — quota volume et quota fichiers par règle
- Suppression optionnelle des **dossiers vides** après purge
- Intégration au **journal d'événements Windows** (journal Application)
- **Codes de sortie** normalisés pour le monitoring de la tâche planifiée

---

## Prérequis

| Élément | Minimum |
|---|---|
| PowerShell | 5.1 (Windows) ou 7+ |
| OS | Windows Server 2016+ / Windows 10+ |
| Droits | Lecture sur les chemins cibles, Écriture pour la suppression |
| `-Parallel` | PowerShell 7+ uniquement |
| Journal Windows | Droits admin requis pour créer une source d'événement |

---

## Installation

```powershell
Copy-Item Invoke-FilePurge.ps1 C:\Scripts\
Unblock-File -Path C:\Scripts\Invoke-FilePurge.ps1
```

---

## Paramètres

### Principaux

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-ConfigFile` | `string` | — | Chemin vers le fichier JSON de règles (mode JSON) |
| `-TargetPath` | `string[]` | — | Un ou plusieurs chemins à purger (mode CLI) |
| `-AgeDays` | `int` | `90` | Âge minimum des fichiers en jours |
| `-UseCreationTime` | `switch` | — | Utilise `CreationTime` au lieu de `LastWriteTime` |
| `-IncludeExtensions` | `string[]` | `@()` (toutes) | Extensions à inclure |
| `-ExcludeExtensions` | `string[]` | `@()` (aucune) | Extensions à exclure |
| `-ExcludePatterns` | `string[]` | `@()` (aucun) | Regex d'exclusion sur le chemin complet |
| `-MaxDeleteMB` | `long` | `10240` | Volume max supprimé par exécution en Mo |
| `-MaxFiles` | `long` | `500000` | Nombre max de fichiers supprimés par exécution |
| `-PurgeEmptyFolders` | `switch` | — | Supprime les dossiers vides après purge |

### Globaux

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-LogPath` | `string` | Dossier du script | Dossier de destination des logs |
| `-LogRetentionDays` | `int` | `30` | Durée de conservation des logs en jours |
| `-WriteEventLog` | `switch` | — | Écrit dans le journal Windows Application |
| `-EventSource` | `string` | `FilePurge` | Nom de la source d'événement Windows |
| `-WhatIf` | `switch` | — | Simulation — aucune suppression |

### Parallélisme (PS7+ uniquement)

| Paramètre | Type | Défaut | Description |
|---|---|---|---|
| `-Parallel` | `switch` | — | Traite les règles en parallèle |
| `-ThrottleLimit` | `int` | `4` | Nombre max de règles simultanées (1–32) |
| `-OldestFirst` | `switch` | — | Trie les candidats du plus ancien au plus récent avant suppression (tri externe, O(1) RAM) |
| `-LogEachFile` | `switch` | — | Logue chaque fichier supprimé (désactivé par défaut sur les gros volumes) |

---

## Configuration JSON

Le fichier JSON contient deux sections : `global` (valeurs partagées) et `rules` (surcharges par chemin).

**Ordre de priorité :** paramètre CLI > champ de la règle > section global > valeur par défaut du code

```json
{
  "global": {
    "LogPath"          : "C:\\Admin\\Logs\\Purge",
    "LogRetentionDays" : 30,
    "WriteEventLog"    : false,
    "MaxDeleteMB"      : 10240,
    "MaxFiles"         : 500000,
    "PurgeEmptyFolders": false
  },
  "rules": [
    {
      "Path"              : "C:\\inetpub\\logs\\LogFiles",
      "AgeDays"           : 90,
      "IncludeExtensions" : [".log"]
    },
    {
      "Path"                : "E:\\Interfaces\\HL7\\Archive",
      "AgeDays"             : 30,
      "IncludeNamePatterns" : ["^Old_[0-9]{2}-[0-9]{2}-[0-9]{4}.*[.]txt$"]
    }
  ]
}
```

### Champs disponibles par règle

| Champ | Type | Description |
|---|---|---|
| `Path` | string | **Obligatoire.** Chemin racine à purger |
| `AgeDays` | int | Âge minimum des fichiers en jours |
| `UseCreationTime` | bool | Utilise `CreationTime` au lieu de `LastWriteTime` |
| `IncludeExtensions` | string[] | Extensions à inclure |
| `ExcludeExtensions` | string[] | Extensions à exclure |
| `IncludeNamePatterns` | string[] | Regex sur le **nom de fichier** — logique OU |
| `IncludePathPatterns` | string[] | Regex sur le **chemin complet** — logique OU |
| `ExcludePatterns` | string[] | Regex sur le **chemin complet** — logique ET (aucune ne doit matcher) |
| `MaxDeleteMB` | long | Quota volume en Mo par règle |
| `MaxFiles` | long | Quota fichiers par règle |
| `PurgeEmptyFolders` | bool | Supprime les dossiers vides après purge |

### Règles de compatibilité JSON PS5.1

PS5.1 utilise `JavaScriptSerializer`, plus strict que PS7 :

| ❌ Interdit | ✅ À utiliser |
|---|---|
| Virgule après le dernier champ | Pas de virgule finale |
| `\d`, `\.` dans les regex | `[0-9]`, `[.]` |
| Clés commençant par `_` dans les tableaux | `Description`, `Note` |
| Commentaires `//` | Pas de commentaires dans le JSON |
| BOM UTF-8 | UTF-8 sans BOM |

---

## Exemples

### 1. Simulation — toujours commencer par là

```powershell
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -WhatIf
```

### 2. Mode JSON — exécution séquentielle

```powershell
.\Invoke-FilePurge.ps1 `
    -ConfigFile "C:\Scripts\purge-rules.json" `
    -LogPath "C:\Admin\Logs" `
    -WriteEventLog
```

### 3. Mode JSON — exécution parallèle (PS7+)

Traite toutes les règles en parallèle, 3 simultanément au maximum.

```powershell
.\Invoke-FilePurge.ps1 `
    -ConfigFile "C:\Scripts\purge-rules.json" `
    -Parallel `
    -ThrottleLimit 3 `
    -WriteEventLog
```

### 4. Mode CLI — logs IIS (rétrocompatible v2.x)

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "C:\inetpub\logs\LogFiles" `
    -AgeDays 90 `
    -IncludeExtensions '.log' `
    -MaxDeleteMB 20480 `
    -LogPath "C:\Admin\Logs"
```

### 5. Tâche planifiée Windows — configuration recommandée

| Champ | Valeur |
|---|---|
| Programme | `powershell.exe` |
| Arguments | voir ci-dessous |
| Démarrer dans | `C:\Scripts` |

**Séquentiel (PS5.1 / PS7) :**
```
powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -ConfigFile "C:\Scripts\purge-rules.json" -WriteEventLog
```

**Parallèle (PS7 uniquement) :**
```
pwsh.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -ConfigFile "C:\Scripts\purge-rules.json" -Parallel -ThrottleLimit 4 -WriteEventLog
```

---

## Traitement parallèle

`-Parallel` utilise `ForEach-Object -Parallel` de PS7 pour traiter les règles simultanément.

- Chaque règle s'exécute dans un **runspace isolé** — les fonctions sont injectées via `${function:X}` / `$using:`
- Les écritures dans le fichier de log sont protégées par un **mutex nommé** — pas d'entrelacement entre règles
- La sortie console (`Write-Host`) est **nativement thread-safe** dans PS7
- Bascule automatiquement en mode séquentiel sur PS5.1 avec un avertissement

**Quand utiliser le parallélisme :**
- Plusieurs règles sur des **volumes physiques différents** — concurrence I/O maximale
- Règles avec une **énumération lourde** (millions de fichiers chacune)

**Quand l'éviter :**
- Règles sur le **même disque** — l'I/O parallèle sur une seule broche est plus lente que séquentielle
- `ThrottleLimit` supérieur au nombre de disques physiques — aucun bénéfice, contention accrue

---

## Modes de suppression

### Streaming (défaut) — recommandé pour les gros volumes

Suppression inline pendant l'énumération — aucune collection en mémoire.

```powershell
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json"
```

| Propriété | Valeur |
|---|---|
| Mémoire | O(1) — compteurs seulement |
| Ordre | Ordre filesystem (NTFS) |
| Risque OOM | Aucun |
| Log progression | Tous les 100 000 fichiers |

### OldestFirst — tri externe, O(1) RAM

Les candidats sont écrits dans des fichiers temporaires triés (chunks de 100 000 enregistrements), puis fusionnés et supprimés du plus ancien au plus récent. Utile quand l'ordre de suppression compte (ex. avec un quota volume).

```powershell
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -OldestFirst
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -OldestFirst -MaxDeleteMB 51200
```

| Propriété | Valeur |
|---|---|
| Mémoire | ~17 Mo par chunk (fixe, jamais croissant) |
| Ordre | Plus ancien en premier (garanti) |
| Risque OOM | Aucun |
| Fichiers temp | Écrits dans `LogPath`, nettoyés dans `finally` |

**Quand utiliser `-OldestFirst`** : quand un quota volume (`-MaxDeleteMB`) est actif et que vous voulez conserver les fichiers les plus récents. Sans quota, streaming et OldestFirst donnent le même résultat.

### Log par fichier

Par défaut, les suppressions individuelles ne sont **pas** loguées pour éviter des logs de plusieurs Go et l'overhead I/O sur les gros volumes. La progression reste toujours loguée tous les 100 000 fichiers.

```powershell
# Réactiver le log par fichier
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -LogEachFile
```

Le rapport CSV contient toujours la liste complète des fichiers supprimés, indépendamment de `-LogEachFile`.

## Codes de sortie

Les messages du log sont en **anglais**.

| Code | Signification | Action recommandée |
|---|---|---|
| `0` | Succès complet | — |
| `1` | Erreur critique (chemin invalide, permissions, JSON) | Vérifier le log |
| `2` | Quota atteint, purge partielle | Augmenter le quota ou planifier plus fréquemment |
| `3` | Avertissement — erreurs sur des fichiers individuels | Consulter le log |

---

## Fichiers générés

| Fichier | Description |
|---|---|
| `FilePurge_YYYYMMDD_HHMMSS.log` | Log structuré complet de l'exécution |
| `FilePurge_YYYYMMDD_HHMMSS_report.csv` | Rapport CSV des fichiers supprimés |

### Niveaux de log

| Icône | Niveau | Signification |
|---|---|---|
| ` ` | `INFO` | Information générale |
| `...` | `DEBUG` | Détail technique |
| `[+]` | `SUCCESS` | Fichier ou dossier supprimé |
| `[!]` | `WARN` | Avertissement (quota, erreur non bloquante) |
| `[X]` | `ERROR` | Erreur bloquante |
| `===` | `SECTION` | Séparateur de section |

---

## Changelog

Voir [CHANGELOG.md](CHANGELOG.md) pour l'historique complet des versions.

---

## Licence

[MIT](LICENSE) — © 2026 9 Lives IT Solutions

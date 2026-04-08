# Invoke-FilePurge.ps1

> 🇫🇷 [Version française disponible](README.fr.md)

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![Version](https://img.shields.io/badge/Version-2.5.0-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

Production-grade PowerShell script for **automated file purging** on Windows environments. Designed for scheduled tasks and large-scale file systems (tested on 8M+ files). Compatible with PS 5.1 and PS 7+.

---

## Features

- Age-based filtering (`LastWriteTime` or `CreationTime`), extension filters, regex exclusion patterns
- **Simulation mode** (`-WhatIf`) — lists candidates without deleting anything
- Timestamped structured log with **automatic rotation**
- **CSV report** of deleted files (path, age, size, status)
- **Circuit breaker** — configurable volume and file count quota per run
- Optional **empty folder cleanup** after purge
- **Windows Event Log** integration (Application log)
- Normalized **exit codes** for Task Scheduler monitoring
- Native .NET enumeration engine — tested on **8M+ files**

---

## Requirements

| Item | Minimum |
|---|---|
| PowerShell | 5.1 (Windows) or 7+ |
| OS | Windows Server 2016+ / Windows 10+ |
| Permissions | Read on target path, Write for deletion |
| Windows Event Log | Admin rights required to register an event source |

---

## Installation

```powershell
# Copy the script to your scripts folder
Copy-Item Invoke-FilePurge.ps1 C:\Scripts\

# Unblock if downloaded from the network
Unblock-File -Path C:\Scripts\Invoke-FilePurge.ps1
```

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-TargetPath` | `string[]` | **Required** | One or more root paths to purge |
| `-AgeDays` | `int` | `90` | Minimum file age in days |
| `-UseCreationTime` | `switch` | — | Use `CreationTime` instead of `LastWriteTime` |
| `-IncludeExtensions` | `string[]` | `@()` (all) | Extensions to include (e.g. `.log`, `.tmp`) |
| `-ExcludeExtensions` | `string[]` | `@()` (none) | Extensions to explicitly exclude |
| `-ExcludePatterns` | `string[]` | `@()` (none) | Regex patterns applied to the full file path |
| `-MaxDeleteMB` | `long` | `10240` | Maximum volume deleted per run in MB (10 GB) |
| `-MaxFiles` | `long` | `500000` | Maximum number of files deleted per run |
| `-LogPath` | `string` | Script folder | Destination folder for log files |
| `-LogRetentionDays` | `int` | `30` | Log retention period in days |
| `-PurgeEmptyFolders` | `switch` | — | Remove empty folders after file purge |
| `-WhatIf` | `switch` | — | Simulation mode — no files are deleted |
| `-WriteEventLog` | `switch` | — | Write an event to the Windows Application log |
| `-EventSource` | `string` | `FilePurge` | Event source name in the Windows log |

---

## Examples

### 1. Simulation — always check before deleting

Run `-WhatIf` first to validate candidates without touching any files.
The `[DEBUG] Diagnostic LastWriteTime` log line shows the oldest/newest dates found and the applied cutoff.

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Logs" `
    -AgeDays 90 `
    -WhatIf
```

---

### 2. Basic purge — IIS logs older than 90 days

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "C:\inetpub\logs\LogFiles" `
    -AgeDays 90 `
    -IncludeExtensions '.log' `
    -LogPath "C:\Admin\Logs"
```

---

### 3. Multi-path purge with exclusions

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Logs", "E:\Temp", "F:\Archives\Import" `
    -AgeDays 60 `
    -IncludeExtensions '.log', '.tmp', '.bak' `
    -ExcludePatterns 'KEEP_', '_PERMANENT', '\\audit\\' `
    -LogPath "D:\Admin\Purge\Logs" `
    -WriteEventLog
```

> `-ExcludePatterns` are **regular expressions** applied to the full file path.
> Examples: `'\\audit\\'` excludes any file under an `audit` folder, `'KEEP_'` excludes files whose path contains `KEEP_`.

---

### 4. Large-scale catch-up purge with empty folder cleanup

Typical use case: initial cleanup on a volume with years of accumulated files.

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

### 5. CreationTime-based purge

Useful when files are regularly copied (resetting `LastWriteTime`) but the creation date remains reliable.

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "E:\Exports\Daily" `
    -AgeDays 30 `
    -UseCreationTime `
    -IncludeExtensions '.csv', '.xlsx' `
    -LogPath "C:\Admin\Logs"
```

---

### 6. Windows Task Scheduler — recommended setup

**Task action configuration:**

| Field | Value |
|---|---|
| Program | `powershell.exe` |
| Arguments | see below |
| Start in | `C:\Scripts` |

```
powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -TargetPath "D:\Logs" -AgeDays 90 -MaxDeleteMB 20480 -LogPath "C:\Admin\Logs" -WriteEventLog
```

**Monitor the exit code** in your alerting tool:

```powershell
$result = Start-Process powershell.exe -ArgumentList '...' -Wait -PassThru
switch ($result.ExitCode) {
    0 { Write-Host "OK — purge completed" }
    1 { Send-Alert "CRITICAL: purge failed (invalid path or permissions)" }
    2 { Send-Alert "WARNING: quota reached, partial purge" }
    3 { Send-Alert "WARNING: errors on some files, check the log" }
}
```

---

## Exit Codes

| Code | Meaning | Recommended action |
|---|---|---|
| `0` | Full success | — |
| `1` | Critical error (invalid path, permissions) | Check the log, fix permissions |
| `2` | Quota reached, partial purge | Increase `-MaxDeleteMB` or schedule more frequently |
| `3` | Warning — errors on individual files | Review the log for per-file errors |

---

## Output Files

| File | Description |
|---|---|
| `FilePurge_YYYYMMDD_HHMMSS.log` | Full structured execution log |
| `FilePurge_YYYYMMDD_HHMMSS_report.csv` | CSV report of deleted files (path, age, size, status) |

### Log format

```
2026-01-15 03:00:01 === [SECTION] ===============================================================
2026-01-15 03:00:01 === [SECTION] INVOKE-FILEPURGE v2.5.0  —  REAL MODE
2026-01-15 03:00:01     [INFO]    Minimum age : 90 days (LastWriteTime) — cutoff : 2025-10-17
2026-01-15 03:00:01 ... [DEBUG]   Enumeration engine : manual .NET Framework recursion (PS5.1)
2026-01-15 03:00:45 ... [DEBUG]   Diagnostic LastWriteTime — oldest : 2023-04-02 | newest : 2026-01-14 | cutoff : 2025-10-17
2026-01-15 03:01:12 [+] [SUCCESS] Deleted : D:\Logs\app_20230402.log  (age: 653d, 2.14 MB)
2026-01-15 03:02:00     [INFO]    Scanned : 45,231 — Candidates : 12,847 — Deleted : 12,847
```

### CSV report format

```csv
"Path","AgeDays","SizeBytes","DeletedAt","Status"
"D:\Logs\app_20230402.log","653","2244608","2026-01-15T03:01:12","Deleted"
"D:\Logs\app_20230403.log","652","1887232","2026-01-15T03:01:12","Deleted"
"D:\Temp\import_20230101.tmp","379","512","2026-01-15T03:01:13","Error: Access denied"
```

---

## Quota Behavior

The quota is a **safety circuit breaker**, not a target. When reached:

- The purge stops immediately — oldest files are always processed first (sorted by date ascending)
- Exit code is set to `2`
- A `[WARN] Quota reached` line appears in the log
- A `Warning` Windows event is emitted if `-WriteEventLog` is active

For initial catch-up on a heavily loaded volume, either increase `-MaxDeleteMB` temporarily or schedule multiple consecutive runs.

---

## PS 5.1 vs PS 7+ Compatibility

The script automatically detects the runtime and selects the optimal enumeration engine:

| Runtime | Engine | Notes |
|---|---|---|
| PS 7+ / .NET 5+ | `EnumerationOptions` | Fastest, native `IgnoreInaccessible` support |
| PS 5.1 / .NET Framework 4.x | Manual `TopDirectoryOnly` recursion | Robust against inaccessible folders |

The `[DEBUG] Enumeration engine:` line in the log confirms which engine is active.

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full version history.

---

## License

[MIT](LICENSE) — © 2026 9 Lives IT Solutions
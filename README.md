# Invoke-FilePurge.ps1

> 🇫🇷 [Version française disponible](README.fr.md)

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue?logo=powershell)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![Version](https://img.shields.io/badge/Version-3.3.4-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

Production-grade PowerShell script for **automated file purging** on Windows environments. JSON-driven per-path rule engine, parallel processing (PS7+), Task Scheduler ready. Compatible with PS 5.1 and PS 7+.

---

## Features

- **JSON rule engine** — per-path filters, quotas and options in a single config file
- **Backward-compatible CLI mode** — all v2.x parameters still work
- **Parallel processing** (`-Parallel`) — concurrent rule execution on PS7+, thread-safe logging
- Age-based filtering (`LastWriteTime` or `CreationTime`), extension filters, regex patterns
  - `IncludeNamePatterns` — regex on **filename**
  - `IncludePathPatterns` — regex on **full path**
  - `ExcludePatterns` — regex on **full path**
- **Simulation mode** (`-WhatIf`) — lists candidates without deleting anything
- Structured timestamped log with **automatic rotation**
- **CSV report** of deleted files (path, age, size, status)
- **Circuit breaker** — per-rule volume and file count quota
- Optional **empty folder cleanup** after purge
- **Windows Event Log** integration (Application log)
- Normalized **exit codes** for Task Scheduler monitoring

---

## Requirements

| Item | Minimum |
|---|---|
| PowerShell | 5.1 (Windows) or 7+ |
| OS | Windows Server 2016+ / Windows 10+ |
| Permissions | Read on target paths, Write for deletion |
| `-Parallel` | PowerShell 7+ only |
| Windows Event Log | Admin rights required to register event source |

---

## Installation

```powershell
Copy-Item Invoke-FilePurge.ps1 C:\Scripts\
Unblock-File -Path C:\Scripts\Invoke-FilePurge.ps1
```

---

## Parameters

### Core

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ConfigFile` | `string` | — | Path to JSON rules file (JSON mode) |
| `-TargetPath` | `string[]` | — | One or more paths to purge (CLI mode) |
| `-AgeDays` | `int` | `90` | Minimum file age in days |
| `-UseCreationTime` | `switch` | — | Use `CreationTime` instead of `LastWriteTime` |
| `-IncludeExtensions` | `string[]` | `@()` (all) | Extensions to include |
| `-ExcludeExtensions` | `string[]` | `@()` (none) | Extensions to exclude |
| `-ExcludePatterns` | `string[]` | `@()` (none) | Regex exclusions on full path |
| `-MaxDeleteMB` | `long` | `10240` | Max volume deleted per run in MB |
| `-MaxFiles` | `long` | `500000` | Max files deleted per run |
| `-PurgeEmptyFolders` | `switch` | — | Remove empty folders after purge |

### Global

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-LogPath` | `string` | Script folder | Log destination folder |
| `-LogRetentionDays` | `int` | `30` | Log retention in days |
| `-WriteEventLog` | `switch` | — | Write to Windows Application log |
| `-EventSource` | `string` | `FilePurge` | Windows event source name |
| `-WhatIf` | `switch` | — | Simulation — no files deleted |

### Parallel (PS7+ only)

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-Parallel` | `switch` | — | Process rules concurrently |
| `-ThrottleLimit` | `int` | `4` | Max concurrent rules (1–32) |

### Deletion behavior (PS 5.1 and PS 7+)

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-OldestFirst` | `switch` | — | External sort — delete oldest files first. O(1) RAM via 100k-record temp chunks. |
| `-LogEachFile` | `switch` | — | Log every deleted file. Disabled by default on large volumes to avoid multi-GB logs. |

---

## JSON Configuration

The JSON file has two sections: `global` (shared defaults) and `rules` (per-path overrides).

**Resolution order:** CLI flag > rule field > global field > code default

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

### Rule fields

| Field | Type | Description |
|---|---|---|
| `Path` | string | **Required.** Root path to purge |
| `AgeDays` | int | Minimum file age in days |
| `UseCreationTime` | bool | Use `CreationTime` instead of `LastWriteTime` |
| `IncludeExtensions` | string[] | Extensions to include |
| `ExcludeExtensions` | string[] | Extensions to exclude |
| `IncludeNamePatterns` | string[] | Regex on **filename** — OR logic |
| `IncludePathPatterns` | string[] | Regex on **full path** — OR logic |
| `ExcludePatterns` | string[] | Regex on **full path** — AND logic (none must match) |
| `MaxDeleteMB` | long | Per-rule volume quota in MB |
| `MaxFiles` | long | Per-rule file count quota |
| `PurgeEmptyFolders` | bool | Remove empty folders after purge |

### PS5.1 JSON compatibility rules

PS5.1 uses `JavaScriptSerializer` which is stricter than PS7:

| ❌ Not allowed | ✅ Use instead |
|---|---|
| Trailing comma on last field | No trailing comma |
| `\d`, `\.` in regex strings | `[0-9]`, `[.]` |
| Keys starting with `_` in arrays | `Description`, `Note` |
| `// comments` | No comments in JSON |
| UTF-8 BOM | UTF-8 without BOM |

---

## Examples

### 1. Simulation — always check first

```powershell
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -WhatIf
```

### 2. JSON mode — sequential

```powershell
.\Invoke-FilePurge.ps1 `
    -ConfigFile "C:\Scripts\purge-rules.json" `
    -LogPath "C:\Admin\Logs" `
    -WriteEventLog
```

### 3. JSON mode — parallel (PS7+)

Processes all rules concurrently, up to 3 at a time. Useful when rules target different volumes.

```powershell
.\Invoke-FilePurge.ps1 `
    -ConfigFile "C:\Scripts\purge-rules.json" `
    -Parallel `
    -ThrottleLimit 3 `
    -WriteEventLog
```

### 4. CLI mode — IIS logs (v2.x compatible)

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "C:\inetpub\logs\LogFiles" `
    -AgeDays 90 `
    -IncludeExtensions '.log' `
    -MaxDeleteMB 20480 `
    -LogPath "C:\Admin\Logs"
```

### 5. CLI mode — multi-path with exclusions

```powershell
.\Invoke-FilePurge.ps1 `
    -TargetPath "D:\Logs", "E:\Temp", "F:\Archives" `
    -AgeDays 60 `
    -IncludeExtensions '.log', '.tmp', '.bak' `
    -ExcludePatterns 'KEEP_', '\\audit\\' `
    -LogPath "C:\Admin\Logs" `
    -WriteEventLog
```

### 6. Windows Task Scheduler — recommended setup

| Field | Value |
|---|---|
| Program | `powershell.exe` |
| Arguments | see below |
| Start in | `C:\Scripts` |

**Sequential (PS5.1 / PS7):**
```
powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -ConfigFile "C:\Scripts\purge-rules.json" -WriteEventLog
```

**Parallel (PS7 only):**
```
pwsh.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass -File "C:\Scripts\Invoke-FilePurge.ps1" -ConfigFile "C:\Scripts\purge-rules.json" -Parallel -ThrottleLimit 4 -WriteEventLog
```

**Monitor the exit code:**

```powershell
$result = Start-Process powershell.exe -ArgumentList '...' -Wait -PassThru
switch ($result.ExitCode) {
    0 { Write-Host "OK" }
    1 { Send-Alert "CRITICAL: purge failed" }
    2 { Send-Alert "WARNING: quota reached, partial purge" }
    3 { Send-Alert "WARNING: errors on some files" }
}
```

---

## Parallel processing

`-Parallel` uses PS7's `ForEach-Object -Parallel` to process rules concurrently.

- Each rule runs in an **isolated runspace** — functions are injected via `${function:X}` / `$using:`
- Log file writes are protected by a **named mutex** — no interleaving between rules
- Console output (`Write-Host`) is **natively thread-safe** in PS7
- Falls back to sequential automatically on PS5.1 with a warning

**When to use parallel:**
- Multiple rules on **different physical volumes** — maximum I/O concurrency
- Rules with **heavy enumeration** (millions of files each)

**When to avoid:**
- Rules on the **same disk** — parallel I/O on a single spindle is slower than sequential
- `-ThrottleLimit` > number of physical drives — no benefit, increased contention

---

## Deletion modes

### Streaming (default) — recommended for large volumes

Files are deleted inline during enumeration. No in-memory collection.

```powershell
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json"
```

| Property | Value |
|---|---|
| Memory | O(1) — counters only |
| Deletion order | Filesystem (NTFS) order |
| OOM risk | None |
| Progress log | Every 100,000 files |

### OldestFirst — external sort, O(1) RAM — PS 5.1 and PS 7+

Candidates are written to sorted temp chunks (100,000 records each), then merged and deleted oldest-first. Useful when deletion order matters (e.g. with a volume quota). Compatible with both PS 5.1 and PS 7+.

```powershell
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -OldestFirst
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -OldestFirst -MaxDeleteMB 51200
```

| Property | Value |
|---|---|
| Memory | ~17 MB per chunk (fixed, never grows) |
| Deletion order | Oldest first (guaranteed) |
| OOM risk | None |
| Temp files | Written to `LogPath`, cleaned up in `finally` |

**When to use `-OldestFirst`:** when a volume quota (`-MaxDeleteMB`) is active and you want to keep the most recent files. Without a quota, streaming and OldestFirst produce the same end result.

### Per-file logging

By default, individual deletions are **not** logged to avoid multi-GB log files and I/O overhead on large volumes. Progress is always logged every 100,000 files.

```powershell
# Re-enable per-file logging
.\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -LogEachFile
```

The CSV report always contains the full list of deleted files regardless of `-LogEachFile`.


---

## Exit Codes

| Code | Meaning | Recommended action |
|---|---|---|
| `0` | Full success | — |
| `1` | Critical error (invalid path, permissions, JSON) | Check the log |
| `2` | Quota reached, partial purge | Increase quota or run more frequently |
| `3` | Warning — errors on individual files | Review the log |

---

## Output Files

| File | Description |
|---|---|
| `FilePurge_YYYYMMDD_HHMMSS.log` | Full structured execution log |
| `FilePurge_YYYYMMDD_HHMMSS_report.csv` | CSV of deleted files (path, age, size, status) |

### Log format

```
2026-01-15 03:00:01 === [SECTION] INVOKE-FILEPURGE v3.3.4  --  REAL MODE
2026-01-15 03:00:01     [INFO]    Execution     : Parallel (ThrottleLimit: 4)
2026-01-15 03:00:01     [INFO]    Rules loaded  : 5 path(s) to process
2026-01-15 03:00:01 === [SECTION] RULE: C:\inetpub\logs\LogFiles
2026-01-15 03:00:01     [INFO]      Age         : 90 days (LastWriteTime) -- cutoff: 2025-10-17
2026-01-15 03:00:45 ... [DEBUG]   Enumeration engine: .NET EnumerationOptions (PS7+)
2026-01-15 03:00:46 ... [DEBUG]   100 000 files scanned, 12 847 deleted, 2.31 GB...
2026-01-15 03:01:00     [INFO]    Scan complete: 45 231 scanned, 12 847 candidates, 12 847 deleted, 2.31 GB.
```

> **Note:** `[SUCCESS] Deleted: ...` lines per file require `-LogEachFile`. By default only progress every 100,000 files is logged.

```
```

---

## Changelog

See [CHANGELOG.md](CHANGELOG.md) for the full version history.

---

## License

[MIT](LICENSE) — © 2026 9 Lives IT Solutions

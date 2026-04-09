# Changelog

All notable changes to this project are documented in this file.

---

## [3.3.4] - 2026-04-09

### Fixed
- `$script:ps51_LogEachFile` not defined in the **streaming** PS5.1 init block
- The v3.3.3 fix had added the variable only in the OldestFirst collect block
- Under `Set-StrictMode -Version Latest`, reading an undefined `$script:` variable throws `VariableIsUndefined` -- caught by the deletion `catch` block, logged as ERROR, but file was NOT deleted (Remove-Item had already succeeded)
- Fix: `$script:ps51_LogEachFile = $LogEachFile` added to both PS5.1 init blocks (streaming + OldestFirst)

---

## [3.3.3] - 2026-04-09

### Performance
- `Add-Content` replaced by persistent `[System.IO.StreamWriter]` with `AutoFlush = $true`
  - `Add-Content` opens, writes and closes the file on every call -- on 8M files this is 8M file-open operations
  - `StreamWriter` is opened once at startup and closed in `finally` -- eliminates I/O overhead per line
  - Parallel mode temp log files still use `Add-Content` (each file is short-lived per rule)
- Individual deletion log entries disabled by default (`-LogEachFile` switch required to re-enable)
  - Logging every deleted file on 8M+ candidates produces multi-GB log files and adds one `WriteLine` per file
  - Default: progress logged every 100,000 files + final summary
  - `-LogEachFile`: restores per-file `[SUCCESS] Deleted: ...` log entries

### Added
- `-LogEachFile` switch parameter (CLI + propagated to parallel runspaces)
- `$script:ps51_LogEachFile` state variable for PS5.1 recursion path

---

## [3.3.2] - 2026-04-09

### Added
- External merge sort for `-OldestFirst` mode -- O(1) memory regardless of candidate count
- `Write-ChunkFile` function: sorts a chunk of 100,000 `{T=ticks, S=size, P=path}` records and writes to a UTF-8 temp file (3 lines per record for `StreamReader` line-by-line reading)
- `Read-ChunkRecord` function: reads one record (3 lines) from an open `StreamReader`
- K-way merge: opens one `StreamReader` per chunk, maintains a head record per reader, selects the minimum ticks (oldest date) in O(k) linear scan, advances the winning reader

### Changed
- `-OldestFirst` no longer collects all candidates in RAM before sorting
- Memory profile: `chunk_size x ~175 bytes ≈ 17 MB` per active chunk, plus `k x ~4 KB` StreamReader buffers during merge
- OOM is no longer possible with `-OldestFirst`
- Temp chunk files written to `LogPath`, cleaned up in `finally` (guaranteed even on error)

---

## [3.3.1] - 2026-04-09

### Added
- `-OldestFirst` switch: collect all candidates in memory, sort by date ascending, delete oldest first
- `$capturedOldestFirst` captured and passed to parallel runspaces
- `-OldestFirst` propagated to `Invoke-PurgeRule` signature

### Note
- This version collected all candidates in RAM -- replaced by external sort in v3.3.2

---

## [3.3.0] - 2026-04-09

### Changed — streaming architecture (O(1) memory)
- `Invoke-PurgeRule` no longer collects `List<FileInfo>` before deleting
- Files are deleted inline during enumeration -- memory is O(1) regardless of candidate count
- `Remove-FilesSafe` function removed (deletion logic inlined into enumeration loop)
- `Invoke-RecurseEnum51` rewritten to perform inline delete during recursion
- WhatIf mode: streams and counts, logs only the first 1,000 matches + summary total
- Progress log every 100,000 files: `N files scanned, N deleted, X GB...`

### Fixed
- `OutOfMemoryException` on 8M+ candidates: v3.2.x collected all `FileInfo` objects before deleting (~2-3 GB RAM for 8M files)

### Added
- New `$script:ps51_*` state variables for inline delete in PS5.1 path: `ps51_IsWhatIf`, `ps51_MaxBytes`, `ps51_StartTime`, `ps51_Report`

---

## [3.2.2] - 2026-04-09

### Fixed
- `$rules` variable referenced before assignment -- premature `$rules.Count` check placed before `Resolve-PurgeRules` call, crashing under `Set-StrictMode -Version Latest`

---

## [3.2.1] - 2026-04-09

### Fixed
- `HasInclExt = ($rInclExt.Count -gt 0)` -- `$rInclExt` can be `$null` in PS5.1 after `Normalize-Extensions`
- All five boolean flags now computed via `Test-RuleArray` (null-safe)

---

## [3.2.0] - 2026-04-09

### Performance
- `Test-RuleArray` rewritten without allocation: `foreach` early-return instead of `List<string>` + `ToArray()`
- `Test-ShouldInclude` uses pre-computed boolean flags (`HasInclExt`, `HasExclExt`, `HasInclName`, `HasInclPath`, `HasExclPat`) -- no function call per file
- PS7+ enumeration: `DirectoryInfo.EnumerateFiles` replaces `Directory.EnumerateFiles` + `new FileInfo(path)` -- OS metadata pre-populated from `FindFirst`/`FindNext`

---

## [3.1.6] - 2026-04-09

### Fixed
- `-Parallel` with a single rule: overhead with zero benefit -- auto-downgrade to sequential with `[WARN]`

---

## [3.1.5] - 2026-04-09

### Fixed
- `finally` called `Dispose()` before last `Write-Log` -- `WaitOne` on disposed handle threw `ObjectDisposedException`
- `Write-Log` mutex acquisition made defensive (`try/catch [ObjectDisposedException]`)

---

## [3.1.4] - 2026-04-09

### Fixed
- `ForEach-Object -Parallel` rejects `ScriptBlock` via `$using:` -- replaced with string-serialized function definitions reconstructed via `[scriptblock]::Create()` inside each runspace
- Per-rule temp log files in parallel mode -- no shared file, no cross-runspace mutex
- `Write-EventLog` absent from PS7 -- replaced by `New-WinEvent` with silent fallback
- Safe `Dispose()` of per-runspace mutex in parallel `finally`

---

## [3.1.3] - 2026-04-09

### Fixed — root cause of 0 candidates (PS5.1)
- `$null | ForEach-Object { ".$_" }` executes once with `$_ = $null` → `Normalize-Extensions @()` produced `['.']` instead of `[]`
- `.xml -notin ['.']` → `$true` → every file rejected
- `Safe-FilterArray` and rewritten `Normalize-Extensions` use explicit `foreach` loops with null-guards

---

## [3.1.2] - 2026-04-09

### Fixed
- `@($null).Count` returns `1` in PS5.1 -- `IncludeNamePatterns` wrongly activated when empty
- `Test-RuleArray` replaces all `.Count` guards in `Test-ShouldInclude`

---

## [3.1.1] - 2026-04-09

### Fixed
- `$function:Write-Log = ...` -- hyphenated function names require `${function:Name}` syntax as assignment target in parallel runspace

---

## [3.1.0] - 2026-04-09

### Added
- `-Parallel` / `-ThrottleLimit` — concurrent rule processing (PS7+ only)
- Named mutex for thread-safe log writes
- `Write-Log -File` explicit parameter for parallel safety

---

## [3.0.2] - 2026-04-08

### Fixed
- PS5.1 array unwrap in `PSCustomObject` properties
- `$1` literal artifact from regex substitution in `.Count` guards

---

## [3.0.1] - 2026-04-08

### Fixed
- `Get-JsonArray` forced `[string[]]` on `rules` array -- each rule coerced to string
- `return @($rules)` -- PS5.1 unwraps empty `List<T>` to `$null`

---

## [3.0.0] - 2026-04-08

### Added
- JSON rule engine (`-ConfigFile`), `global` section, per-rule fields
- `IncludeNamePatterns` (filename regex, OR logic) and `IncludePathPatterns` (full path regex, OR logic)
- Resolution order: CLI > rule > global > default
- JSON parse error hints (trailing comma, underscore keys, BOM detection)

---

## [2.5.0] - 2026-04-08

### Fixed
- Root cause of 0 candidates: empty pipeline returned `$null` in PS5.1, causing `PropertyNotFoundException` on `.Count` under StrictMode, silently swallowed by `catch {}`

---

## [2.4.0] - 2026-04-08

### Fixed
- `Invoke-RecurseEnum51` defined in `else {}` block scope instead of script scope

---

## [2.3.0] - 2026-04-08

### Fixed
- `List<FileInfo>` copied by PS5.1 binding engine when passed as typed parameter -- migrated to `$script:ps51_*` state

---

## [2.2.0] - 2026-04-08

### Fixed
- PS5.1 variable scope in recursive function

---

## [2.1.0] - 2026-04-08

### Fixed
- `EnumerationOptions` absent from .NET Framework 4.x -- dual-engine with runtime detection

---

## [2.0.0] - 2026-04-08

### Added
- Initial production release: dual enumeration engine, structured log, CSV report, WhatIf, Windows Event Log, normalized exit codes

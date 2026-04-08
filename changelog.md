# Changelog

All notable changes to this project will be documented in this file.

## [2.5.0] - 2026-04-08
### Fixed
- `$null.Count` under `Set-StrictMode -Version Latest` — empty pipeline now
  guaranteed as `[string[]]@()` instead of `$null`
- Generic `catch {}` replaced by typed IO catches — unexpected logic exceptions
  now logged at DEBUG level instead of being silently swallowed

## [2.4.0] - 2026-04-08
### Fixed
- `Invoke-RecurseEnum51` was defined inside `else {}` block — moved to script
  scope to guarantee visibility across all loop iterations

## [2.3.0] - 2026-04-08
### Added
- Date diagnostic logged after each scan (oldest / newest / cutoff)
- Cutoff date displayed in log header
### Fixed
- Mutable state in PS5.1 path migrated to `$script:ps51_*` variables —
  `List<FileInfo>` passed as typed parameter was silently copied by PS5.1 binding

## [2.2.0] - 2026-04-08
### Added
- Explicit parameters in recursive function
- Cutoff date logged in header
### Fixed
- PS5.1 variable scope issues in recursive enumeration

## [2.1.0] - 2026-04-08
### Fixed
- `EnumerationOptions` not available in .NET Framework 4.x (PS5.1)
- Replaced with manual `TopDirectoryOnly` recursion — robust to inaccessible folders

## [2.0.0] - 2026-04-08
### Added
- Initial production release
- Dual-engine enumeration: `EnumerationOptions` (PS7+) / manual recursion (PS5.1)
- Structured log with automatic rotation
- CSV report of deleted files
- Volume and file count quota (circuit breaker)
- WhatIf simulation mode
- Windows Event Log integration
- Normalized exit codes for Task Scheduler monitoring
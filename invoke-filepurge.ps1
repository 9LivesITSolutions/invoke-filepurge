#Requires -Version 5.1
<#
.SYNOPSIS
    Automated purge of old files with structured logging, reporting and advanced error handling.

.DESCRIPTION
    Production-grade script designed for Windows Task Scheduler.
    Features:
      - Age-based filtering (LastWriteTime or CreationTime), extension filters, exclusion patterns
      - Simulation mode (WhatIf) — no files deleted
      - Structured timestamped log with automatic rotation
      - Summary report (console + file + optional Windows Event Log)
      - Optional recursive empty folder cleanup
      - Volume quota (circuit breaker) per run
      - Normalized exit codes for Task Scheduler monitoring
      - Compatible with PS 5.1+ and PS 7+

.PARAMETER TargetPath
    One or more root paths to purge. Accepts an array of paths.

.PARAMETER AgeDays
    Minimum file age in days (based on LastWriteTime). Default: 90.

.PARAMETER UseCreationTime
    If specified, uses CreationTime instead of LastWriteTime for age calculation.

.PARAMETER IncludeExtensions
    List of extensions to include (e.g. '.log', '.tmp', '.bak'). Empty = all extensions.

.PARAMETER ExcludeExtensions
    List of extensions to explicitly exclude.

.PARAMETER ExcludePatterns
    Regular expressions applied to the full file path for exclusion.

.PARAMETER MaxDeleteMB
    Maximum volume deleted per run in MB (safety circuit breaker). Default: 10240 (10 GB).

.PARAMETER MaxFiles
    Maximum number of files deleted per run. Default: 500000.

.PARAMETER LogPath
    Destination folder for log files. Default: script folder.

.PARAMETER LogRetentionDays
    Log retention period in days. Default: 30.

.PARAMETER PurgeEmptyFolders
    If specified, removes empty folders after file purge.

.PARAMETER WhatIf
    Simulation mode — lists files that would be deleted without touching anything.

.PARAMETER WriteEventLog
    Writes an event to the Windows Application log at the end of execution.

.PARAMETER EventSource
    Windows event source name. Default: "FilePurge".

.EXAMPLE
    # Simulation — list .log files older than 60 days
    .\Invoke-FilePurge.ps1 -TargetPath "D:\Logs" -AgeDays 60 -IncludeExtensions '.log' -WhatIf

.EXAMPLE
    # Real purge — multiple paths, exclusions, 5 GB quota, log to D:\Purge\Logs
    .\Invoke-FilePurge.ps1 `
        -TargetPath "D:\Logs","E:\Archives\Temp" `
        -AgeDays 365 `
        -IncludeExtensions '.log','.tmp','.bak' `
        -ExcludePatterns 'KEEP_','_PERMANENT' `
        -MaxDeleteMB 5120 `
        -LogPath "D:\Purge\Logs" `
        -PurgeEmptyFolders `
        -WriteEventLog

.EXAMPLE
    # Task Scheduler — minimal production call
    powershell.exe -NonInteractive -NoProfile -ExecutionPolicy Bypass `
        -File "C:\Scripts\Invoke-FilePurge.ps1" `
        -TargetPath "C:\inetpub\logs\LogFiles" `
        -AgeDays 90 `
        -MaxDeleteMB 20480 `
        -WriteEventLog

.NOTES
    Author        : 9 Lives IT Solutions — production-grade PowerShell
    Version       : 2.5.0
    Compatibility : Windows Server 2016+ / PS 5.1 / PS 7+
    Exit codes    :
        0  = Success (no files found or purge completed)
        1  = Critical error (invalid path, permissions, etc.)
        2  = Quota reached (MaxDeleteMB or MaxFiles exceeded, partial purge)
        3  = Warning (errors on some files, partial purge)
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string[]] $TargetPath,

    [ValidateRange(0, 36500)]
    [int] $AgeDays = 90,

    [switch] $UseCreationTime,

    [string[]] $IncludeExtensions = @(),

    [string[]] $ExcludeExtensions = @(),

    [string[]] $ExcludePatterns = @(),

    [ValidateRange(1, 1048576)]
    [long] $MaxDeleteMB = 10240,

    [ValidateRange(1, 10000000)]
    [long] $MaxFiles = 500000,

    [string] $LogPath = '',

    [ValidateRange(1, 365)]
    [int] $LogRetentionDays = 30,

    [switch] $PurgeEmptyFolders,

    [switch] $WriteEventLog,

    [string] $EventSource = 'FilePurge'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
# REGION : Initialization
# -----------------------------------------------------------------------------
#region Init

$Script:Version   = '2.5.0'
$Script:StartTime = Get-Date
$Script:IsWhatIf  = $WhatIfPreference -or ($PSBoundParameters.ContainsKey('WhatIf'))
$Script:ExitCode  = 0

# Normalize extensions to lowercase with leading dot.
# @(...) wrapper is MANDATORY: in PS5.1, an empty pipeline returns $null (not @()).
# Set-StrictMode -Version Latest raises PropertyNotFoundException on $null.Count,
# which is silently swallowed by catch{} -- resulting in 0 candidates. This wrap
# guarantees a [string[]] array even when no extensions are provided.
[string[]] $NormInclude = @($IncludeExtensions | ForEach-Object { if ($_ -notmatch '^\.' ) { ".$_" } else { $_ } } | ForEach-Object { $_.ToLower() })
[string[]] $NormExclude = @($ExcludeExtensions | ForEach-Object { if ($_ -notmatch '^\.' ) { ".$_" } else { $_ } } | ForEach-Object { $_.ToLower() })

# -- Log folder ----------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Split-Path -Parent $PSCommandPath
}
if (-not (Test-Path $LogPath)) {
    New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
}

$Timestamp  = $Script:StartTime.ToString('yyyyMMdd_HHmmss')
$LogFile    = Join-Path $LogPath "FilePurge_$Timestamp.log"
$ReportFile = Join-Path $LogPath "FilePurge_$Timestamp`_report.csv"

# -- Counters ------------------------------------------------------------------
$Stats = [PSCustomObject]@{
    FilesScanned   = [long] 0
    FilesMatched   = [long] 0
    FilesDeleted   = [long] 0
    FilesSkipped   = [long] 0
    FilesErrored   = [long] 0
    BytesDeleted   = [long] 0
    FoldersDeleted = [long] 0
    QuotaReached   = $false
}

$Report = [System.Collections.Generic.List[PSCustomObject]]::new()

#endregion

# -----------------------------------------------------------------------------
# REGION : Helper functions
# -----------------------------------------------------------------------------
#region Functions

function Write-Log {
    param (
        [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG','SECTION')]
        [string] $Level = 'INFO'
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $icon = switch ($Level) {
        'INFO'    { '   ' }
        'WARN'    { '[!]' }
        'ERROR'   { '[X]' }
        'SUCCESS' { '[+]' }
        'DEBUG'   { '...' }
        'SECTION' { '===' }
    }
    $line = "$ts $icon [$Level] $Message"

    # Colored console output
    $color = switch ($Level) {
        'INFO'    { 'Cyan'    }
        'WARN'    { 'Yellow'  }
        'ERROR'   { 'Red'     }
        'SUCCESS' { 'Green'   }
        'DEBUG'   { 'Gray'    }
        'SECTION' { 'Magenta' }
        default   { 'White'   }
    }
    Write-Host $line -ForegroundColor $color

    # Append to log file
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
}

function Format-Bytes {
    param ([long] $Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-FileAge {
    param ([System.IO.FileInfo] $File)
    $refDate = if ($UseCreationTime) { $File.CreationTime } else { $File.LastWriteTime }
    return (New-TimeSpan -Start $refDate -End $Script:StartTime).TotalDays
}

function Test-ShouldInclude {
    param ([System.IO.FileInfo] $File)
    $ext = $File.Extension.ToLower()

    # Include filter
    if ($NormInclude.Count -gt 0 -and $ext -notin $NormInclude) { return $false }

    # Exclude extension filter
    if ($NormExclude.Count -gt 0 -and $ext -in $NormExclude) { return $false }

    # Regex exclusion patterns applied to full path
    foreach ($pattern in $ExcludePatterns) {
        if ($File.FullName -match $pattern) { return $false }
    }

    return $true
}

function Remove-FilesSafe {
    param (
        [System.IO.FileInfo[]] $Files,
        [long] $MaxBytes
    )

    $deletedBytes = [long] 0

    foreach ($file in $Files) {

        # -- Quota check -------------------------------------------------------
        if ($Stats.FilesDeleted -ge $MaxFiles) {
            Write-Log "File quota reached ($MaxFiles). Stopping early." -Level WARN
            $Stats.QuotaReached = $true
            $Script:ExitCode = 2
            break
        }
        if ($deletedBytes + $file.Length -gt $MaxBytes) {
            Write-Log "Volume quota reached ($MaxDeleteMB MB). Stopping early." -Level WARN
            $Stats.QuotaReached = $true
            $Script:ExitCode = 2
            break
        }

        # -- Deletion ----------------------------------------------------------
        $age = [math]::Round((Get-FileAge $file), 1)

        if ($Script:IsWhatIf) {
            Write-Log "[WHATIF] Would delete: $($file.FullName)  (age: ${age}d, $(Format-Bytes $file.Length))" -Level DEBUG
            $Stats.FilesDeleted++
            $Stats.BytesDeleted += $file.Length
        }
        else {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                Write-Log "Deleted: $($file.FullName)  (age: ${age}d, $(Format-Bytes $file.Length))" -Level SUCCESS
                $Stats.FilesDeleted++
                $Stats.BytesDeleted += $file.Length
                $deletedBytes       += $file.Length

                $Report.Add([PSCustomObject]@{
                    Path       = $file.FullName
                    AgeDays    = $age
                    SizeBytes  = $file.Length
                    DeletedAt  = (Get-Date -Format 'o')
                    Status     = 'Deleted'
                })
            }
            catch {
                Write-Log "ERROR deleting: $($file.FullName) -- $($_.Exception.Message)" -Level ERROR
                $Stats.FilesErrored++
                if ($Script:ExitCode -lt 3) { $Script:ExitCode = 3 }
                $Report.Add([PSCustomObject]@{
                    Path       = $file.FullName
                    AgeDays    = $age
                    SizeBytes  = $file.Length
                    DeletedAt  = ''
                    Status     = "Error: $($_.Exception.Message)"
                })
            }
        }
    }
}

function Remove-EmptyFolders {
    param ([string] $RootPath)
    Write-Log "Scanning for empty folders under: $RootPath" -Level INFO

    # Reverse sort by path length (leaves first) for cascading cleanup
    $emptyFolders = Get-ChildItem -LiteralPath $RootPath -Recurse -Directory -ErrorAction SilentlyContinue |
        Where-Object { (Get-ChildItem -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0 } |
        Sort-Object { $_.FullName.Length } -Descending

    foreach ($folder in $emptyFolders) {
        if ($Script:IsWhatIf) {
            Write-Log "[WHATIF] Would remove empty folder: $($folder.FullName)" -Level DEBUG
            $Stats.FoldersDeleted++
        }
        else {
            try {
                Remove-Item -LiteralPath $folder.FullName -Force -ErrorAction Stop
                Write-Log "Empty folder removed: $($folder.FullName)" -Level SUCCESS
                $Stats.FoldersDeleted++
            }
            catch {
                Write-Log "ERROR removing empty folder: $($folder.FullName) -- $($_.Exception.Message)" -Level WARN
            }
        }
    }
}

function Invoke-LogRotation {
    Write-Log "Rotating logs (retention: $LogRetentionDays days)..." -Level INFO
    $limit = $Script:StartTime.AddDays(-$LogRetentionDays)
    $oldLogs = Get-ChildItem -Path $LogPath -File -Filter 'FilePurge_*' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limit }
    foreach ($old in $oldLogs) {
        try {
            Remove-Item -LiteralPath $old.FullName -Force
            Write-Log "Old log removed: $($old.Name)" -Level DEBUG
        }
        catch {
            Write-Log "Could not remove log: $($old.Name)" -Level WARN
        }
    }
}

function Invoke-RecurseEnum51 {
    # -- Manual recursion for PS5.1 / .NET Framework ---------------------------
    # All mutable state is accessed via $script:ps51_* variables -- never as
    # typed generic parameters (PS5.1 binding engine would copy the collection,
    # making internal Add() calls invisible to the caller).
    # Only immutable scalar passed as parameter: the current directory path.
    param ([string] $DirPath)

    # -- Files in current folder -----------------------------------------------
    try {
        $files = [System.IO.Directory]::EnumerateFiles(
                     $DirPath, '*', [System.IO.SearchOption]::TopDirectoryOnly)

        foreach ($filePath in $files) {
            $script:ps51_Stats.FilesScanned++
            try {
                $fi = [System.IO.FileInfo]::new($filePath)

                # Skip junctions and symlinks
                if ($fi.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }

                $refDate = if ($script:ps51_UseCreation) { $fi.CreationTime } else { $fi.LastWriteTime }

                # Date diagnostic tracking (independent of purge filter)
                if ($refDate -lt $script:ps51_DiagOldest) { $script:ps51_DiagOldest = $refDate }
                if ($refDate -gt $script:ps51_DiagNewest) { $script:ps51_DiagNewest = $refDate }

                # Purge condition
                if ($refDate -lt $script:ps51_Cutoff -and (Test-ShouldInclude $fi)) {
                    $script:ps51_Files.Add($fi)   # direct mutation on the original List<>
                    $script:ps51_Stats.FilesMatched++
                }
            }
            catch [System.IO.IOException]       { <# File locked / inaccessible -- skipped #> }
            catch [UnauthorizedAccessException] { <# Access denied -- skipped #> }
            catch {
                # Unexpected exception (StrictMode, cast, logic error) -- logged for debugging
                Write-Log "Warning on '$filePath': $($_.Exception.GetType().Name) -- $($_.Exception.Message)" -Level DEBUG
            }

            if ($script:ps51_Stats.FilesScanned % 100000 -eq 0) {
                Write-Log "$($script:ps51_Stats.FilesScanned.ToString('N0')) files scanned, $($script:ps51_Stats.FilesMatched.ToString('N0')) candidates..." -Level DEBUG
            }
        }
    }
    catch [UnauthorizedAccessException] { <# Protected folder -- skipped #> }
    catch { <# Other access error -- skipped #> }

    # -- Subfolders (recursion) ------------------------------------------------
    try {
        $subDirs = [System.IO.Directory]::EnumerateDirectories(
                       $DirPath, '*', [System.IO.SearchOption]::TopDirectoryOnly)

        foreach ($sub in $subDirs) {
            try {
                $di = [System.IO.DirectoryInfo]::new($sub)
                if ($di.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                Invoke-RecurseEnum51 -DirPath $sub
            }
            catch { <# Inaccessible subfolder -- skipped #> }
        }
    }
    catch [UnauthorizedAccessException] { <# Directory listing denied -- skipped #> }
    catch { <# Other error #> }
}

function Write-EventLogEntry {
    param ([string] $Message, [string] $EntryType = 'Information', [int] $EventId = 1000)
    # NOTE: [EventLog]::SourceExists() tries to read the Security log and throws
    # if the process lacks sufficient admin rights.
    # Strategy: attempt New-EventLog directly (idempotent if source already exists),
    # then write. Each call is isolated in its own try/catch.
    try {
        New-EventLog -LogName Application -Source $EventSource -ErrorAction SilentlyContinue
    }
    catch {
        # Source already registered or insufficient rights to create it --
        # we still attempt to write; the Application log is generally accessible.
    }
    try {
        Write-EventLog -LogName Application -Source $EventSource -EntryType $EntryType `
            -EventId $EventId -Message $Message
    }
    catch {
        Write-Log "Could not write to Windows Event Log: $($_.Exception.Message)" -Level WARN
    }
}

#endregion

# -----------------------------------------------------------------------------
# REGION : Main execution
# -----------------------------------------------------------------------------
#region Main

try {

    # -- Header ----------------------------------------------------------------
    Write-Log ('=' * 70) -Level SECTION
    Write-Log "INVOKE-FILEPURGE v$Script:Version  --  $(if ($Script:IsWhatIf) { 'SIMULATION MODE (WhatIf)' } else { 'REAL MODE' })" -Level SECTION
    Write-Log ('=' * 70) -Level SECTION
    Write-Log "Started         : $($Script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Log "Target path(s)  : $($TargetPath -join ' | ')"
    Write-Log "Minimum age     : $AgeDays days ($( if ($UseCreationTime) { 'CreationTime' } else { 'LastWriteTime' } )) -- cutoff: $($Script:StartTime.AddDays(-$AgeDays).ToString('yyyy-MM-dd'))"
    Write-Log "Include ext.    : $(if ($NormInclude) { $NormInclude -join ', ' } else { '(all)' })"
    Write-Log "Exclude ext.    : $(if ($NormExclude) { $NormExclude -join ', ' } else { '(none)' })"
    Write-Log "Exclude patterns: $(if ($ExcludePatterns) { $ExcludePatterns -join ' | ' } else { '(none)' })"
    Write-Log "Volume quota    : $(Format-Bytes ($MaxDeleteMB * 1MB))"
    Write-Log "File quota      : $MaxFiles"
    Write-Log "Empty folders   : $(if ($PurgeEmptyFolders) { 'YES' } else { 'NO' })"
    Write-Log "Log file        : $LogFile"
    Write-Log ('─' * 70) -Level SECTION

    $MaxBytes   = $MaxDeleteMB * 1MB
    $CutoffDate = $Script:StartTime.AddDays(-$AgeDays)

    # -- Loop over each target path --------------------------------------------
    foreach ($root in $TargetPath) {

        Write-Log "PROCESSING: $root" -Level SECTION

        # Validate path -- Test-Path can throw on some network paths or restrictive
        # ACLs under PS5.1 instead of returning $false
        $rootAccessible = $false
        try {
            $rootAccessible = Test-Path -LiteralPath $root -PathType Container -ErrorAction Stop
        }
        catch {
            Write-Log "Error accessing path '$root': $($_.Exception.Message)" -Level ERROR
            $Script:ExitCode = 1
            continue
        }
        if (-not $rootAccessible) {
            Write-Log "Path not found or not accessible: $root" -Level ERROR
            $Script:ExitCode = 1
            continue
        }

        # -- Enumeration -------------------------------------------------------
        Write-Log "Enumerating files (large volumes -- please wait)..." -Level INFO

        $allFiles = [System.Collections.Generic.List[System.IO.FileInfo]]::new()

        try {
            # -- Runtime detection: EnumerationOptions only available on .NET 5+ (PS7+) --
            $useModernEnum = $PSVersionTable.PSVersion.Major -ge 7

            if ($useModernEnum) {
                # -- PS 7+ / .NET 5+: EnumerationOptions (fastest) ----------------------
                Write-Log "Enumeration engine: .NET EnumerationOptions (PS7+ / .NET5+)" -Level DEBUG
                $enumOptions = [System.IO.EnumerationOptions]::new()
                $enumOptions.RecurseSubdirectories = $true
                $enumOptions.IgnoreInaccessible    = $true
                $enumOptions.AttributesToSkip      = [System.IO.FileAttributes]::System -bor
                                                     [System.IO.FileAttributes]::ReparsePoint

                $rawFiles = [System.IO.Directory]::EnumerateFiles($root, '*', $enumOptions)

                foreach ($filePath in $rawFiles) {
                    $Stats.FilesScanned++
                    try {
                        $fi = [System.IO.FileInfo]::new($filePath)
                        $refDate = if ($UseCreationTime) { $fi.CreationTime } else { $fi.LastWriteTime }
                        if ($refDate -lt $CutoffDate -and (Test-ShouldInclude $fi)) {
                            $allFiles.Add($fi)
                            $Stats.FilesMatched++
                        }
                    }
                    catch { <# Inaccessible file -- silently skipped #> }

                    if ($Stats.FilesScanned % 100000 -eq 0) {
                        Write-Log "$($Stats.FilesScanned.ToString('N0')) files scanned, $($Stats.FilesMatched.ToString('N0')) candidates..." -Level DEBUG
                    }
                }
            }
            else {
                # -- PS 5.1 / .NET Framework 4.x: robust manual recursion ---------------
                # Core PS5.1 rule: complex generic types (List<T>) passed as typed
                # parameters are COPIED by the binding engine -- internal Add() calls
                # do not affect the original collection.
                # Solution: mutable state in $script:ps51_*, parameters only for
                # immutable scalar values.
                Write-Log "Enumeration engine: manual .NET Framework recursion (PS5.1)" -Level DEBUG

                # -- Shared state in script scope (never pass ref types as parameters) ---
                $script:ps51_Files       = $allFiles       # List<FileInfo> -- mutated directly
                $script:ps51_Stats       = $Stats          # PSCustomObject -- mutated directly
                $script:ps51_Cutoff      = $CutoffDate     # DateTime snapshot
                $script:ps51_UseCreation = [bool]$UseCreationTime
                $script:ps51_DiagOldest  = [DateTime]::MaxValue
                $script:ps51_DiagNewest  = [DateTime]::MinValue

                Invoke-RecurseEnum51 -DirPath $root

                # $allFiles already points to the same object as $script:ps51_Files -- no resync needed

                # -- Date diagnostic report ----------------------------------------
                if ($Stats.FilesScanned -gt 0) {
                    $dateField = if ($script:ps51_UseCreation) { 'CreationTime' } else { 'LastWriteTime' }
                    $oldestStr = if ($script:ps51_DiagOldest -eq [DateTime]::MaxValue) { 'N/A' } else { $script:ps51_DiagOldest.ToString('yyyy-MM-dd') }
                    $newestStr = if ($script:ps51_DiagNewest -eq [DateTime]::MinValue) { 'N/A' } else { $script:ps51_DiagNewest.ToString('yyyy-MM-dd') }
                    Write-Log "Diagnostic $dateField -- oldest: $oldestStr  |  newest: $newestStr  |  cutoff: $($CutoffDate.ToString('yyyy-MM-dd'))" -Level DEBUG

                    if ($Stats.FilesMatched -eq 0 -and $script:ps51_DiagOldest -gt $CutoffDate) {
                        Write-Log "DIAGNOSTIC: No file is older than the cutoff ($($CutoffDate.ToString('yyyy-MM-dd'))). Oldest file found: $oldestStr. Increase -AgeDays or use -UseCreationTime." -Level WARN
                    }
                }
            }
        }
        catch [UnauthorizedAccessException] {
            Write-Log "Access denied to root path: $root" -Level ERROR
            $Script:ExitCode = 1
            continue
        }
        catch {
            Write-Log "Enumeration error on $root : $($_.Exception.Message)" -Level ERROR
            $Script:ExitCode = 1
            continue
        }

        Write-Log "Scan complete: $($Stats.FilesScanned.ToString('N0')) files, $($Stats.FilesMatched.ToString('N0')) candidates for purge." -Level INFO

        if ($allFiles.Count -eq 0) {
            Write-Log "No eligible files under $root." -Level INFO
            continue
        }

        # Sort by date ascending (oldest first)
        $sortedFiles = $allFiles | Sort-Object { if ($UseCreationTime) { $_.CreationTime } else { $_.LastWriteTime } }

        # -- Deletion ----------------------------------------------------------
        Remove-FilesSafe -Files $sortedFiles -MaxBytes $MaxBytes

        if ($Stats.QuotaReached) {
            Write-Log "Quota reached -- processing stopped for $root." -Level WARN
            break
        }

        # -- Empty folders -----------------------------------------------------
        if ($PurgeEmptyFolders) {
            Remove-EmptyFolders -RootPath $root
        }
    }

    # -- CSV report export -----------------------------------------------------
    if ($Report.Count -gt 0 -and -not $Script:IsWhatIf) {
        $Report | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report exported: $ReportFile" -Level INFO
    }

    # -- Log rotation ----------------------------------------------------------
    Invoke-LogRotation

    # -- Summary ---------------------------------------------------------------
    $Duration = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
    $Summary = @"
EXECUTION SUMMARY
  Mode           : $(if ($Script:IsWhatIf) { 'SIMULATION' } else { 'REAL' })
  Duration       : $("{0:hh\:mm\:ss}" -f $Duration)
  Files scanned  : $($Stats.FilesScanned.ToString('N0'))
  Candidates     : $($Stats.FilesMatched.ToString('N0'))
  Deleted        : $($Stats.FilesDeleted.ToString('N0'))
  Volume deleted : $(Format-Bytes $Stats.BytesDeleted)
  Errors         : $($Stats.FilesErrored)
  Empty folders  : $($Stats.FoldersDeleted)
  Quota reached  : $(if ($Stats.QuotaReached) { 'YES (partial purge)' } else { 'No' })
  Exit code      : $Script:ExitCode
"@

    Write-Log ('=' * 70) -Level SECTION
    $Summary.Split("`n") | ForEach-Object { Write-Log $_ -Level $(if ($_ -match 'Error|YES|SIMULATION') { 'WARN' } else { 'INFO' }) }
    Write-Log ('=' * 70) -Level SECTION

    # -- Windows Event Log -----------------------------------------------------
    if ($WriteEventLog) {
        $entryType = switch ($Script:ExitCode) {
            0 { 'Information' }
            1 { 'Error' }
            2 { 'Warning' }
            3 { 'Warning' }
        }
        Write-EventLogEntry -Message $Summary -EntryType $entryType -EventId (1000 + $Script:ExitCode)
    }

}
catch {
    Write-Log "UNHANDLED CRITICAL ERROR: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    $Script:ExitCode = 1
    if ($WriteEventLog) {
        Write-EventLogEntry -Message "CRITICAL ERROR FilePurge: $($_.Exception.Message)" -EntryType Error -EventId 1099
    }
}
finally {
    Write-Log "Full log: $LogFile" -Level INFO
    exit $Script:ExitCode
}

#endregion
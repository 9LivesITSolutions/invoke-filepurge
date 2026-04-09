#Requires -Version 5.1
<#
.SYNOPSIS
    Automated file purge -- JSON rule engine, parallel processing (PS7+), Task Scheduler ready.

.DESCRIPTION
    Production-grade script designed for Windows Task Scheduler.

    Modes:
      - CLI mode  : backward-compatible with v2.x, all parameters on the command line
      - JSON mode : -ConfigFile points to a rules file with per-path filters and quotas

    Features:
      - Per-path rules: AgeDays, extensions, IncludeNamePatterns (filename regex),
        IncludePathPatterns (full path regex), ExcludePatterns, quota, empty folder cleanup
      - JSON global section for shared defaults; CLI flags override the global section
      - Parallel rule processing (PS7+ only) via -Parallel / -ThrottleLimit
        Thread-safe console output; per-rule log segments merged in order after completion
      - WhatIf simulation, structured log with rotation, CSV report, Windows Event Log
      - Dual enumeration engine: EnumerationOptions (PS7+) / manual recursion (PS5.1)
      - Normalized exit codes for monitoring

.PARAMETER TargetPath
    [CLI mode] One or more root paths to purge.

.PARAMETER AgeDays
    [CLI mode] Minimum file age in days. Default: 90.

.PARAMETER UseCreationTime
    Use CreationTime instead of LastWriteTime.

.PARAMETER IncludeExtensions
    [CLI mode] Extensions to include (e.g. '.log','.tmp'). Empty = all.

.PARAMETER ExcludeExtensions
    [CLI mode] Extensions to explicitly exclude.

.PARAMETER ExcludePatterns
    [CLI mode] Regex patterns applied to the full file path for exclusion.

.PARAMETER MaxDeleteMB
    Maximum volume deleted per run in MB. Default: 10240. Overrides JSON if specified.

.PARAMETER MaxFiles
    Maximum files deleted per run. Default: 500000. Overrides JSON if specified.

.PARAMETER PurgeEmptyFolders
    [CLI mode] Remove empty folders after purge.

.PARAMETER ConfigFile
    [JSON mode] Path to the JSON rules file.

.PARAMETER LogPath
    Destination folder for logs. Overrides JSON global.LogPath if specified.

.PARAMETER LogRetentionDays
    Log retention in days. Default: 30. Overrides JSON global.LogRetentionDays.

.PARAMETER WriteEventLog
    Write a Windows Application event at end of run.

.PARAMETER EventSource
    Windows event source name. Default: "FilePurge".

.PARAMETER Parallel
    [PS7+ only] Process rules concurrently. Silently ignored on PS5.1.
    Per-rule console output is real-time; log file segments are flushed after each job.

.PARAMETER ThrottleLimit
    [PS7+ only] Max concurrent rules when -Parallel is active. Default: 4.

.PARAMETER OldestFirst
    Sort candidates by date (oldest first) before deleting.
    Requires collecting all candidates in memory -- may cause OutOfMemoryException
    on very large volumes (millions of candidates). Use only when deletion order
    matters and available RAM is sufficient. Default: streaming (filesystem order).

.EXAMPLE
    # JSON mode -- sequential
    .\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -WhatIf

.EXAMPLE
    # JSON mode -- parallel (PS7+), 3 rules processed concurrently
    .\Invoke-FilePurge.ps1 -ConfigFile "C:\Scripts\purge-rules.json" -Parallel -ThrottleLimit 3

.EXAMPLE
    # CLI mode (v2.x compatible)
    .\Invoke-FilePurge.ps1 -TargetPath "D:\Logs" -AgeDays 90 -IncludeExtensions '.log' -WhatIf

.NOTES
    Author        : 9 Lives IT Solutions -- production-grade PowerShell
    Version       : 3.3.2
    Compatibility : Windows Server 2016+ / PS 5.1 / PS 7+
    Exit codes    :
        0  = Success
        1  = Critical error (invalid path, permissions, JSON parse failure)
        2  = Quota reached (MaxDeleteMB or MaxFiles exceeded, partial purge)
        3  = Warning (errors on some individual files)
#>

[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param (
    # -- CLI mode (v2.x compatible) -------------------------------------------
    [string[]] $TargetPath,

    [ValidateRange(0, 36500)]
    [int] $AgeDays = 90,

    [switch] $UseCreationTime,

    [string[]] $IncludeExtensions = @(),
    [string[]] $ExcludeExtensions = @(),
    [string[]] $ExcludePatterns   = @(),

    [ValidateRange(1, 1048576)]
    [long] $MaxDeleteMB = 10240,

    [ValidateRange(1, 10000000)]
    [long] $MaxFiles = 500000,

    [switch] $PurgeEmptyFolders,

    # -- JSON mode ------------------------------------------------------------
    [string] $ConfigFile = '',

    # -- Global parameters (CLI overrides JSON global section) ----------------
    [string] $LogPath = '',

    [ValidateRange(1, 365)]
    [int] $LogRetentionDays = 30,

    [switch] $WriteEventLog,

    [string] $EventSource = 'FilePurge',

    # -- Parallel (PS7+ only) -------------------------------------------------
    [switch] $Parallel,

    [ValidateRange(1, 32)]
    [int] $ThrottleLimit = 4,

    # -- Sorted deletion (uses more memory) -----------------------------------
    [switch] $OldestFirst
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# =============================================================================
# REGION : Initialization
# =============================================================================
#region Init

$Script:Version   = '3.3.2'
$Script:StartTime = Get-Date
$Script:IsWhatIf  = $WhatIfPreference -or ($PSBoundParameters.ContainsKey('WhatIf'))
$Script:ExitCode  = 0

# Thread-safe log mutex (named per PID -- avoids collision between concurrent script instances)
$Script:LogMutex = [System.Threading.Mutex]::new($false, "FilePurge_$PID")

# -- Early JSON load ----------------------------------------------------------
$script:jsonContent = $null
$script:jsonGlobal  = $null

if (-not [string]::IsNullOrWhiteSpace($ConfigFile)) {
    if (-not (Test-Path -LiteralPath $ConfigFile -PathType Leaf)) {
        Write-Error "ConfigFile not found: $ConfigFile"
        exit 1
    }
    try {
        $rawJson = [System.IO.File]::ReadAllText((Resolve-Path $ConfigFile).Path)
        if ($rawJson.Length -gt 0 -and [int][char]$rawJson[0] -eq 65279) {
            $rawJson = $rawJson.Substring(1)   # strip BOM
        }
        $script:jsonContent = $rawJson | ConvertFrom-Json
        $gm = $script:jsonContent | Get-Member -Name 'global' -MemberType NoteProperty -ErrorAction SilentlyContinue
        if ($gm) { $script:jsonGlobal = $script:jsonContent.global }
    }
    catch {
        $hint = switch -Regex ($_.Exception.Message) {
            'Primitive'  { 'Check for trailing commas or keys starting with underscore inside arrays.' }
            'Unexpected' { 'Check for missing quotes, extra commas, or malformed brackets.' }
            default      { 'Validate the file at https://jsonlint.com' }
        }
        Write-Error "Failed to parse ConfigFile '$ConfigFile': $($_.Exception.Message)`nHint: $hint"
        exit 1
    }
}

# -- Resolve LogPath: CLI explicit > JSON global > script folder --------------
function Resolve-LogPath {
    if ($PSBoundParameters.ContainsKey('LogPath') -and -not [string]::IsNullOrWhiteSpace($LogPath)) {
        return $LogPath
    }
    if ($null -ne $script:jsonGlobal) {
        $m = $script:jsonGlobal | Get-Member -Name 'LogPath' -MemberType NoteProperty -ErrorAction SilentlyContinue
        if ($m -and -not [string]::IsNullOrWhiteSpace($script:jsonGlobal.LogPath)) {
            return $script:jsonGlobal.LogPath
        }
    }
    return Split-Path -Parent $PSCommandPath
}

$LogPath   = Resolve-LogPath
if (-not (Test-Path $LogPath)) { New-Item -ItemType Directory -Path $LogPath -Force | Out-Null }

$Timestamp  = $Script:StartTime.ToString('yyyyMMdd_HHmmss')
$LogFile    = Join-Path $LogPath "FilePurge_$Timestamp.log"
$ReportFile = Join-Path $LogPath "FilePurge_$Timestamp`_report.csv"

# -- Global counters (merged from all rule results after execution) -----------
$GlobalStats = [PSCustomObject]@{
    FilesScanned   = [long] 0
    FilesMatched   = [long] 0
    FilesDeleted   = [long] 0
    FilesErrored   = [long] 0
    BytesDeleted   = [long] 0
    FoldersDeleted = [long] 0
    QuotaReached   = $false
}

$GlobalReport = [System.Collections.Generic.List[PSCustomObject]]::new()

#endregion

# =============================================================================
# REGION : Helper functions
# =============================================================================
#region Functions

function Write-Log {
    param (
        [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG','SECTION')]
        [string] $Level = 'INFO',
        [string] $File  = $LogFile   # explicit file path for parallel safety
    )
    $ts   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $icon = switch ($Level) {
        'INFO'    { '   ' } 'WARN'    { '[!]' } 'ERROR'   { '[X]' }
        'SUCCESS' { '[+]' } 'DEBUG'   { '...' } 'SECTION' { '===' }
    }
    $line  = "$ts $icon [$Level] $Message"
    $color = switch ($Level) {
        'INFO'    { 'Cyan' }    'WARN'    { 'Yellow' }  'ERROR'   { 'Red' }
        'SUCCESS' { 'Green' }   'DEBUG'   { 'Gray' }    'SECTION' { 'Magenta' }
        default   { 'White' }
    }
    Write-Host $line -ForegroundColor $color   # Write-Host is thread-safe in PS7

    # Mutex-protected file write -- safe for parallel execution
    $mutexAcquired = $false
    try {
        $mutexAcquired = $Script:LogMutex.WaitOne(5000)
    }
    catch [System.ObjectDisposedException] { }
    catch { }
    try { Add-Content -Path $File -Value $line -Encoding UTF8 }
    catch { }
    finally {
        if ($mutexAcquired) {
            try { $Script:LogMutex.ReleaseMutex() } catch { }
        }
    }
}

function Format-Bytes {
    param ([long] $Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-FileAge {
    param ([System.IO.FileInfo] $File, [PSCustomObject] $Rule)
    $refDate = if ($Rule.UseCreationTime) { $File.CreationTime } else { $File.LastWriteTime }
    return (New-TimeSpan -Start $refDate -End $Script:StartTime).TotalDays
}

function Get-JsonProp {
    param ($Obj, [string] $Name, $Default = $null)
    if ($null -eq $Obj) { return $Default }
    $m = $Obj | Get-Member -MemberType NoteProperty -Name $Name -ErrorAction SilentlyContinue
    if ($m) { $v = $Obj.$Name; if ($null -ne $v) { return $v } }
    return $Default
}

function Get-JsonArray {
    param ($Obj, [string] $Name, [string[]] $Default = @())
    $val = Get-JsonProp $Obj $Name $null
    if ($null -eq $val) { return $Default }
    [string[]] @($val)
}

function Safe-FilterArray {
    # PS5.1: $null|Where-Object runs ONCE with $_=$null, producing @($null) (Count=1).
    # Always use this instead of inline | Where-Object for rule array properties.
    param ([string[]] $Arr)
    if ($null -eq $Arr) { return [string[]] @() }
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $Arr) {
        if ($null -ne $item -and $item -ne '') { $list.Add($item) }
    }
    if ($list.Count -eq 0) { return [string[]] @() }
    return [string[]] $list.ToArray()
}

function Normalize-Extensions {
    # PS5.1: $null|ForEach-Object runs ONCE with $_=$null, turning @() into ['.'].
    # Use an explicit foreach loop -- never pipe potentially-null/empty arrays.
    param ([string[]] $Ext)
    if ($null -eq $Ext) { return [string[]] @() }
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($e in $Ext) {
        if ($null -ne $e -and $e -ne '') {
            $norm = if ($e -notmatch '^\.') { ".$e" } else { $e }
            $list.Add($norm.ToLower())
        }
    }
    if ($list.Count -eq 0) { return [string[]] @() }
    return [string[]] $list.ToArray()
}

function Test-RuleArray {
    # PS5.1-safe non-empty check -- zero allocations, returns bool immediately.
    # NEVER allocate (List<T>, @(), Where-Object) here: called once per file per filter.
    param ($Arr)
    if ($null -eq $Arr) { return $false }
    foreach ($item in $Arr) {
        if ($null -ne $item -and $item -ne '') { return $true }
    }
    return $false
}

function Test-ShouldInclude {
    # Hot path: called once per candidate file. Use pre-computed bool flags on $Rule
    # (HasInclExt, HasExclExt, etc.) to avoid any function call or allocation per file.
    param ([System.IO.FileInfo] $File, [PSCustomObject] $Rule)
    $ext = $File.Extension.ToLower()

    if ($Rule.HasInclExt  -and $ext -notin $Rule.IncludeExtensions) { return $false }
    if ($Rule.HasExclExt  -and $ext -in  $Rule.ExcludeExtensions)  { return $false }

    if ($Rule.HasInclName) {
        $hit = $false
        foreach ($p in $Rule.IncludeNamePatterns) { if ($File.Name -match $p) { $hit = $true; break } }
        if (-not $hit) { return $false }
    }
    if ($Rule.HasInclPath) {
        $hit = $false
        foreach ($p in $Rule.IncludePathPatterns) { if ($File.FullName -match $p) { $hit = $true; break } }
        if (-not $hit) { return $false }
    }
    if ($Rule.HasExclPat) {
        foreach ($p in $Rule.ExcludePatterns) {
            if ($null -ne $p -and $p -ne '' -and $File.FullName -match $p) { return $false }
        }
    }
    return $true
}

function Resolve-PurgeRules {
    $rules = [System.Collections.Generic.List[PSCustomObject]]::new()

    if (-not [string]::IsNullOrWhiteSpace($ConfigFile)) {
        $g = $script:jsonGlobal
        $jsonRulesProp = Get-JsonProp $script:jsonContent 'rules' $null
        $jsonRules     = if ($null -ne $jsonRulesProp) { @($jsonRulesProp) } else { @() }

        foreach ($r in $jsonRules) {
            $path = Get-JsonProp $r 'Path' $null
            if ([string]::IsNullOrWhiteSpace($path)) {
                Write-Log "Rule skipped -- missing 'Path' field in config." -Level WARN
                continue
            }
            $rAgeDays     = [int]  (Get-JsonProp $r 'AgeDays'           (Get-JsonProp $g 'AgeDays'           $AgeDays))
            $rUseCreation = [bool] (Get-JsonProp $r 'UseCreationTime'    (Get-JsonProp $g 'UseCreationTime'   ([bool]$UseCreationTime)))
            $rPurgeEmpty  = [bool] (Get-JsonProp $r 'PurgeEmptyFolders'  (Get-JsonProp $g 'PurgeEmptyFolders' ([bool]$PurgeEmptyFolders)))
            $rMaxMB       = [long] (Get-JsonProp $r 'MaxDeleteMB'        (Get-JsonProp $g 'MaxDeleteMB'        $MaxDeleteMB))
            $rMaxFiles    = [long] (Get-JsonProp $r 'MaxFiles'           (Get-JsonProp $g 'MaxFiles'           $MaxFiles))

            if ($PSBoundParameters.ContainsKey('MaxDeleteMB')) { $rMaxMB   = $MaxDeleteMB }
            if ($PSBoundParameters.ContainsKey('MaxFiles'))    { $rMaxFiles = $MaxFiles    }

            $rInclExt  = Normalize-Extensions (Get-JsonArray $r 'IncludeExtensions'   (Get-JsonArray $g 'IncludeExtensions'))
            $rExclExt  = Normalize-Extensions (Get-JsonArray $r 'ExcludeExtensions'   (Get-JsonArray $g 'ExcludeExtensions'))
            $rInclName = Safe-FilterArray (Get-JsonArray $r 'IncludeNamePatterns' (Get-JsonArray $g 'IncludeNamePatterns'))
            $rInclPath = Safe-FilterArray (Get-JsonArray $r 'IncludePathPatterns' (Get-JsonArray $g 'IncludePathPatterns'))
            $rExclPat  = Safe-FilterArray (Get-JsonArray $r 'ExcludePatterns'     (Get-JsonArray $g 'ExcludePatterns'))

            $rules.Add([PSCustomObject]@{
                Path                = $path
                AgeDays             = $rAgeDays
                UseCreationTime     = $rUseCreation
                IncludeExtensions   = Normalize-Extensions $rInclExt
                ExcludeExtensions   = Normalize-Extensions $rExclExt
                IncludeNamePatterns = Safe-FilterArray $rInclName
                IncludePathPatterns = Safe-FilterArray $rInclPath
                ExcludePatterns     = Safe-FilterArray $rExclPat
                MaxDeleteMB         = $rMaxMB
                MaxFiles            = $rMaxFiles
                PurgeEmptyFolders   = $rPurgeEmpty
                HasInclExt          = (Test-RuleArray $rInclExt)
                HasExclExt          = (Test-RuleArray $rExclExt)
                HasInclName         = (Test-RuleArray $rInclName)
                HasInclPath         = (Test-RuleArray $rInclPath)
                HasExclPat          = (Test-RuleArray $rExclPat)
            })
        }
        if ($rules.Count -eq 0) { Write-Log "No valid rules found in ConfigFile." -Level WARN }
    }
    else {
        $normInclude = Normalize-Extensions $IncludeExtensions
        $normExclude = Normalize-Extensions $ExcludeExtensions
        foreach ($tp in $TargetPath) {
            $rules.Add([PSCustomObject]@{
                Path                = $tp
                AgeDays             = $AgeDays
                UseCreationTime     = [bool] $UseCreationTime
                IncludeExtensions   = Normalize-Extensions $normInclude
                ExcludeExtensions   = Normalize-Extensions $normExclude
                IncludeNamePatterns = [string[]] @()
                IncludePathPatterns = [string[]] @()
                ExcludePatterns     = Safe-FilterArray $ExcludePatterns
                MaxDeleteMB         = $MaxDeleteMB
                MaxFiles            = $MaxFiles
                PurgeEmptyFolders   = [bool] $PurgeEmptyFolders
                HasInclExt          = (Test-RuleArray $normInclude)
                HasExclExt          = (Test-RuleArray $normExclude)
                HasInclName         = $false
                HasInclPath         = $false
                HasExclPat          = (Test-RuleArray (Safe-FilterArray $ExcludePatterns))
            })
        }
    }

    # Force array return -- PS5.1 unwraps empty List<T> to $null
    return @($rules)
}


function Write-ChunkFile {
    # Sorts a chunk of {T=ticks,S=size,P=path} records and writes them to a temp file.
    # Format: 3 lines per record (ticks / size / path) for StreamReader line-by-line reading.
    param (
        [System.Collections.Generic.List[psobject]] $Chunk,
        [string] $TempFolder
    )
    $sorted   = $Chunk | Sort-Object { $_.T }
    $filePath = [System.IO.Path]::Combine($TempFolder, "purge_chunk_$([System.Guid]::NewGuid().ToString('N')).tmp")
    $sw = [System.IO.StreamWriter]::new($filePath, $false, [System.Text.Encoding]::UTF8)
    try {
        foreach ($r in $sorted) {
            $sw.WriteLine($r.T)   # DateTime.Ticks as string
            $sw.WriteLine($r.S)   # file size as string
            $sw.WriteLine($r.P)   # full path
        }
    }
    finally { $sw.Dispose() }
    return $filePath
}

function Read-ChunkRecord {
    param ([System.IO.StreamReader] $Reader)
    $t = $Reader.ReadLine()
    if ($null -eq $t) { return $null }
    $s = $Reader.ReadLine()
    $p = $Reader.ReadLine()
    if ($null -eq $s -or $null -eq $p) { return $null }
    return [pscustomobject]@{ T = [long] $t; S = [long] $s; P = $p }
}

function Invoke-RecurseEnum51 {
    # Manual recursion for PS5.1 / .NET Framework.
    # Streaming mode: files are processed (deleted or WhatIf-logged) inline -- no List<FileInfo>.
    # All mutable state via $script:ps51_* to avoid PS5.1 binding engine copy issue.
    param ([string] $DirPath)

    try {
        $files = [System.IO.Directory]::EnumerateFiles($DirPath, '*', [System.IO.SearchOption]::TopDirectoryOnly)
        foreach ($filePath in $files) {
            $script:ps51_Stats.FilesScanned++
            try {
                $fi = [System.IO.FileInfo]::new($filePath)
                if ($fi.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                $refDate = if ($script:ps51_UseCreation) { $fi.CreationTime } else { $fi.LastWriteTime }

                # Diagnostic tracking (independent of filter)
                if ($refDate -lt $script:ps51_DiagOldest) { $script:ps51_DiagOldest = $refDate }
                if ($refDate -gt $script:ps51_DiagNewest) { $script:ps51_DiagNewest = $refDate }

                if ($refDate -lt $script:ps51_Cutoff -and (Test-ShouldInclude $fi $script:ps51_Rule)) {
                    $script:ps51_Stats.FilesMatched++

                    # Quota check before processing
                    if ($script:ps51_Stats.FilesDeleted -ge $script:ps51_Rule.MaxFiles) {
                        Write-Log "File quota reached ($($script:ps51_Rule.MaxFiles)). Stopping." -Level WARN -File $script:ps51_LogFile
                        $script:ps51_Stats.QuotaReached = $true
                        return   # Stop recursion
                    }
                    if ($script:ps51_Stats.BytesDeleted + $fi.Length -gt $script:ps51_MaxBytes) {
                        Write-Log "Volume quota reached ($($script:ps51_Rule.MaxDeleteMB) MB). Stopping." -Level WARN -File $script:ps51_LogFile
                        $script:ps51_Stats.QuotaReached = $true
                        return
                    }

                    $age = [math]::Round(($script:ps51_StartTime - $refDate).TotalDays, 1)

                    if ($script:ps51_IsWhatIf) {
                        # WhatIf: log only first 1000 matches to avoid console flood
                        if ($script:ps51_Stats.FilesMatched -le 1000) {
                            Write-Log "[WHATIF] Would delete: $($fi.FullName)  (age: ${age}d, $(Format-Bytes $fi.Length))" -Level DEBUG -File $script:ps51_LogFile
                        }
                        $script:ps51_Stats.FilesDeleted++
                        $script:ps51_Stats.BytesDeleted += $fi.Length
                    }
                    else {
                        try {
                            Remove-Item -LiteralPath $fi.FullName -Force -ErrorAction Stop
                            Write-Log "Deleted: $($fi.FullName)  (age: ${age}d, $(Format-Bytes $fi.Length))" -Level SUCCESS -File $script:ps51_LogFile
                            $script:ps51_Stats.FilesDeleted++
                            $script:ps51_Stats.BytesDeleted += $fi.Length
                            $script:ps51_Report.Add([PSCustomObject]@{
                                Path      = $fi.FullName
                                AgeDays   = $age
                                SizeBytes = $fi.Length
                                DeletedAt = (Get-Date -Format 'o')
                                Status    = 'Deleted'
                            })
                        }
                        catch {
                            Write-Log "ERROR deleting: $($fi.FullName) -- $($_.Exception.Message)" -Level ERROR -File $script:ps51_LogFile
                            $script:ps51_Stats.FilesErrored++
                            if ($script:ps51_Stats.ExitCode -lt 3) { $script:ps51_Stats.ExitCode = 3 }
                        }
                    }
                }
            }
            catch [System.IO.IOException]       { }
            catch [UnauthorizedAccessException] { }
            catch {
                Write-Log "Warning on '$filePath': $($_.Exception.GetType().Name) -- $($_.Exception.Message)" -Level DEBUG -File $script:ps51_LogFile
            }

            # Stop recursion if quota reached
            if ($script:ps51_Stats.QuotaReached) { return }

            if ($script:ps51_Stats.FilesScanned % 100000 -eq 0) {
                Write-Log "$($script:ps51_Stats.FilesScanned.ToString('N0')) files scanned, $($script:ps51_Stats.FilesDeleted.ToString('N0')) deleted..." -Level DEBUG -File $script:ps51_LogFile
            }
        }
    }
    catch [UnauthorizedAccessException] { }
    catch { }

    # Stop recursion if quota reached
    if ($script:ps51_Stats.QuotaReached) { return }

    try {
        $subDirs = [System.IO.Directory]::EnumerateDirectories($DirPath, '*', [System.IO.SearchOption]::TopDirectoryOnly)
        foreach ($sub in $subDirs) {
            try {
                $di = [System.IO.DirectoryInfo]::new($sub)
                if ($di.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                Invoke-RecurseEnum51 -DirPath $sub
                if ($script:ps51_Stats.QuotaReached) { return }
            }
            catch { }
        }
    }
    catch [UnauthorizedAccessException] { }
    catch { }
}

function Invoke-PurgeRule {
    <#
    .SYNOPSIS
        Process a single purge rule. Returns a result PSCustomObject.
        Streaming architecture: files are deleted inline during enumeration.
        No List<FileInfo> collection -- O(1) memory regardless of candidate count.
        WhatIf mode: streams and counts; logs only the first 1000 matches.
    #>
    param (
        [PSCustomObject] $Rule,
        [string]         $RuleLogFile,
        [bool]           $IsWhatIf,
        [datetime]       $StartTime,
        [bool]           $OldestFirst = $false
    )

    $result = [PSCustomObject]@{
        Path           = $Rule.Path
        ExitCode       = 0
        FilesScanned   = [long] 0
        FilesMatched   = [long] 0
        FilesDeleted   = [long] 0
        FilesErrored   = [long] 0
        BytesDeleted   = [long] 0
        FoldersDeleted = [long] 0
        QuotaReached   = $false
        ReportEntries  = [System.Collections.Generic.List[PSCustomObject]]::new()
    }

    $root       = $Rule.Path
    $cutoffDate = $StartTime.AddDays(-$Rule.AgeDays)
    $maxBytes   = $Rule.MaxDeleteMB * 1MB
    $dateField  = if ($Rule.UseCreationTime) { 'CreationTime' } else { 'LastWriteTime' }

    # -- Rule header ----------------------------------------------------------
    Write-Log ('-' * 70) -Level SECTION -File $RuleLogFile
    Write-Log "RULE: $root" -Level SECTION -File $RuleLogFile
    Write-Log "  Age            : $($Rule.AgeDays) days ($dateField) -- cutoff: $($cutoffDate.ToString('yyyy-MM-dd'))" -File $RuleLogFile
    Write-Log "  Include ext.   : $(if ($Rule.HasInclExt)  { $Rule.IncludeExtensions  -join ', ' } else { '(all)' })" -File $RuleLogFile
    Write-Log "  Exclude ext.   : $(if ($Rule.HasExclExt)  { $Rule.ExcludeExtensions  -join ', ' } else { '(none)' })" -File $RuleLogFile
    Write-Log "  Incl. name rx  : $(if ($Rule.HasInclName) { $Rule.IncludeNamePatterns -join ' | ' } else { '(none)' })" -File $RuleLogFile
    Write-Log "  Incl. path rx  : $(if ($Rule.HasInclPath) { $Rule.IncludePathPatterns -join ' | ' } else { '(none)' })" -File $RuleLogFile
    Write-Log "  Excl. path rx  : $(if ($Rule.HasExclPat)  { $Rule.ExcludePatterns     -join ' | ' } else { '(none)' })" -File $RuleLogFile
    Write-Log "  Volume quota   : $(Format-Bytes $maxBytes)  |  File quota: $($Rule.MaxFiles)" -File $RuleLogFile
    Write-Log "  Empty folders  : $(if ($Rule.PurgeEmptyFolders) { 'YES' } else { 'NO' })" -File $RuleLogFile
    $modeLabel = if ($IsWhatIf) { 'SIMULATION -- first 1000 matches logged' }
                  elseif ($OldestFirst) { 'OLDEST-FIRST -- candidates collected in memory then sorted' }
                  else { 'STREAMING DELETE -- O(1) memory, filesystem order' }
    Write-Log "  Mode           : $modeLabel" -File $RuleLogFile

    # -- Validate path --------------------------------------------------------
    $rootAccessible = $false
    try { $rootAccessible = Test-Path -LiteralPath $root -PathType Container -ErrorAction Stop }
    catch {
        Write-Log "Error accessing path '$root': $($_.Exception.Message)" -Level ERROR -File $RuleLogFile
        $result.ExitCode = 1; return $result
    }
    if (-not $rootAccessible) {
        Write-Log "Path not found or not accessible: $root" -Level ERROR -File $RuleLogFile
        $result.ExitCode = 1; return $result
    }

    # -- OldestFirst: external sort (chunk-based, O(1) RAM) -------------------
    # Phase 1: enumerate -> write sorted chunks of 100 000 records to temp files
    # Phase 2: k-way merge of all chunks -> delete in oldest-first order
    # RAM: max ~17 MB per chunk regardless of total candidate count.
    if ($OldestFirst) {
        $chunkSize  = 100000
        $tempFolder = $LogPath   # reuse log folder for temp chunks
        $chunkFiles = [System.Collections.Generic.List[string]]::new()
        $chunk      = [System.Collections.Generic.List[psobject]]::new($chunkSize)
        $chunkNum   = 0

        Write-Log "Mode: OldestFirst -- external sort (chunks of $($chunkSize.ToString('N0')), temp: $tempFolder)" -File $RuleLogFile

        # Helper: flush current chunk to disk
        $flushChunk = {
            if ($chunk.Count -eq 0) { return }
            $chunkNum++
            $cf = Write-ChunkFile -Chunk $chunk -TempFolder $tempFolder
            $chunkFiles.Add($cf)
            Write-Log "Chunk $chunkNum written: $($chunk.Count.ToString('N0')) records -> $([System.IO.Path]::GetFileName($cf))" -Level DEBUG -File $RuleLogFile
            $chunk.Clear()
        }

        # -- Phase 1: enumerate and chunk -----------------------------------------
        try {
            if ($PSVersionTable.PSVersion.Major -ge 7) {
                $enumOpts2 = [System.IO.EnumerationOptions]::new()
                $enumOpts2.RecurseSubdirectories = $true
                $enumOpts2.IgnoreInaccessible    = $true
                $enumOpts2.AttributesToSkip      = [System.IO.FileAttributes]::System -bor [System.IO.FileAttributes]::ReparsePoint
                foreach ($fi in ([System.IO.DirectoryInfo]::new($root)).EnumerateFiles('*', $enumOpts2)) {
                    $result.FilesScanned++
                    $rd = if ($Rule.UseCreationTime) { $fi.CreationTime } else { $fi.LastWriteTime }
                    if ($rd -lt $cutoffDate -and (Test-ShouldInclude $fi $Rule)) {
                        $result.FilesMatched++
                        $chunk.Add([pscustomobject]@{ T = $rd.Ticks; S = $fi.Length; P = $fi.FullName })
                        if ($chunk.Count -ge $chunkSize) { & $flushChunk }
                    }
                    if ($result.FilesScanned % 100000 -eq 0) {
                        Write-Log "$($result.FilesScanned.ToString('N0')) scanned, $($result.FilesMatched.ToString('N0')) candidates, $chunkNum chunk(s) written..." -Level DEBUG -File $RuleLogFile
                    }
                }
            }
            else {
                # PS5.1 -- streaming collect into chunk via ps51_ state
                $script:ps51_Stats       = $result
                $script:ps51_Cutoff      = $cutoffDate
                $script:ps51_UseCreation = $Rule.UseCreationTime
                $script:ps51_Rule        = $Rule
                $script:ps51_LogFile     = $RuleLogFile
                $script:ps51_IsWhatIf    = $true           # collect-only, no delete
                $script:ps51_MaxBytes    = [long]::MaxValue
                $script:ps51_StartTime   = $StartTime
                $script:ps51_Report      = $result.ReportEntries
                $script:ps51_DiagOldest  = [DateTime]::MaxValue
                $script:ps51_DiagNewest  = [DateTime]::MinValue
                $script:ps51_DiagDone    = $false
                # Override: intercept each match into chunk instead of reporting
                # Use ps51_IsWhatIf=true so Invoke-RecurseEnum51 counts but doesn't delete.
                # After, iterate $result.FilesMatched records -- but we have no list.
                # PS5.1 path: collect to chunk inline by re-implementing scan here.
                $files51 = [System.IO.Directory]::EnumerateFiles($root, '*', [System.IO.SearchOption]::AllDirectories)
                foreach ($fp in $files51) {
                    $result.FilesScanned++
                    try {
                        $fi51 = [System.IO.FileInfo]::new($fp)
                        if ($fi51.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }
                        $rd51 = if ($Rule.UseCreationTime) { $fi51.CreationTime } else { $fi51.LastWriteTime }
                        if ($rd51 -lt $cutoffDate -and (Test-ShouldInclude $fi51 $Rule)) {
                            $result.FilesMatched++
                            $chunk.Add([pscustomobject]@{ T = $rd51.Ticks; S = $fi51.Length; P = $fi51.FullName })
                            if ($chunk.Count -ge $chunkSize) { & $flushChunk }
                        }
                    }
                    catch { }
                    if ($result.FilesScanned % 100000 -eq 0) {
                        Write-Log "$($result.FilesScanned.ToString('N0')) scanned, $($result.FilesMatched.ToString('N0')) candidates..." -Level DEBUG -File $RuleLogFile
                    }
                }
            }
            & $flushChunk   # flush remaining
        }
        catch {
            Write-Log "Enumeration error (OldestFirst): $($_.Exception.Message)" -Level ERROR -File $RuleLogFile
            foreach ($cf in $chunkFiles) { Remove-Item $cf -Force -ErrorAction SilentlyContinue }
            $result.ExitCode = 1; return $result
        }

        Write-Log "Phase 1 complete: $($result.FilesMatched.ToString('N0')) candidates in $($chunkFiles.Count) chunk(s). Starting k-way merge..." -File $RuleLogFile

        # -- Phase 2: k-way merge and inline delete --------------------------------
        $readers = [System.Collections.Generic.List[System.IO.StreamReader]]::new()
        $heads   = [System.Collections.Generic.List[psobject]]::new()
        try {
            foreach ($cf in $chunkFiles) {
                $sr = [System.IO.StreamReader]::new($cf, [System.Text.Encoding]::UTF8)
                $readers.Add($sr)
                $heads.Add((Read-ChunkRecord $sr))
            }

            $mergedCount = [long] 0
            while ($true) {
                # Find chunk with minimum ticks (oldest date) -- linear scan O(k)
                $minTick = [long]::MaxValue
                $minIdx  = -1
                for ($i = 0; $i -lt $heads.Count; $i++) {
                    if ($null -ne $heads[$i] -and $heads[$i].T -lt $minTick) {
                        $minTick = $heads[$i].T
                        $minIdx  = $i
                    }
                }
                if ($minIdx -eq -1) { break }   # all chunks exhausted

                $rec = $heads[$minIdx]
                $mergedCount++

                # Quota check
                if ($result.FilesDeleted -ge $Rule.MaxFiles) {
                    Write-Log "File quota reached ($($Rule.MaxFiles)). Stopping merge." -Level WARN -File $RuleLogFile
                    $result.QuotaReached = $true; break
                }
                if ($result.BytesDeleted + $rec.S -gt $maxBytes) {
                    Write-Log "Volume quota reached ($($Rule.MaxDeleteMB) MB). Stopping merge." -Level WARN -File $RuleLogFile
                    $result.QuotaReached = $true; break
                }

                $age = [math]::Round(($StartTime - [datetime]::new($rec.T)).TotalDays, 1)

                if ($IsWhatIf) {
                    if ($result.FilesDeleted -lt 1000) {
                        Write-Log "[WHATIF] Would delete: $($rec.P)  (age: ${age}d, $(Format-Bytes $rec.S))" -Level DEBUG -File $RuleLogFile
                    }
                    elseif ($result.FilesDeleted -eq 1000) {
                        Write-Log "[WHATIF] ... (1000+ matches -- remaining not listed)" -Level DEBUG -File $RuleLogFile
                    }
                    $result.FilesDeleted++
                    $result.BytesDeleted += $rec.S
                }
                else {
                    try {
                        Remove-Item -LiteralPath $rec.P -Force -ErrorAction Stop
                        Write-Log "Deleted: $($rec.P)  (age: ${age}d, $(Format-Bytes $rec.S))" -Level SUCCESS -File $RuleLogFile
                        $result.FilesDeleted++
                        $result.BytesDeleted += $rec.S
                        $result.ReportEntries.Add([PSCustomObject]@{
                            Path = $rec.P; AgeDays = $age; SizeBytes = $rec.S
                            DeletedAt = (Get-Date -Format 'o'); Status = 'Deleted'
                        })
                    }
                    catch {
                        Write-Log "ERROR: $($rec.P) -- $($_.Exception.Message)" -Level ERROR -File $RuleLogFile
                        $result.FilesErrored++
                        if ($result.ExitCode -lt 3) { $result.ExitCode = 3 }
                    }
                }

                # Advance reader that gave the minimum
                $heads[$minIdx] = Read-ChunkRecord $readers[$minIdx]

                if ($mergedCount % 100000 -eq 0) {
                    Write-Log "$($mergedCount.ToString('N0')) merged, $($result.FilesDeleted.ToString('N0')) deleted, $(Format-Bytes $result.BytesDeleted)..." -Level DEBUG -File $RuleLogFile
                }
            }
        }
        finally {
            foreach ($sr in $readers) { try { $sr.Dispose() } catch { } }
            foreach ($cf in $chunkFiles) { Remove-Item $cf -Force -ErrorAction SilentlyContinue }
        }

        Write-Log "OldestFirst complete: $($result.FilesScanned.ToString('N0')) scanned, $($result.FilesMatched.ToString('N0')) candidates, $($result.FilesDeleted.ToString('N0')) deleted, $(Format-Bytes $result.BytesDeleted)." -File $RuleLogFile
        if ($IsWhatIf -and $result.FilesMatched -gt 1000) {
            Write-Log "[WHATIF] Summary: $($result.FilesMatched.ToString('N0')) files would be deleted ($(Format-Bytes $result.BytesDeleted))." -Level WARN -File $RuleLogFile
        }

        if ($Rule.PurgeEmptyFolders -and -not $result.QuotaReached) {
            $doEmptyFolders = $true
        }
        return $result
    }

    # -- Enumeration + inline processing (streaming) --------------------------
    Write-Log "Enumerating and processing files (streaming -- no in-memory collection)..." -File $RuleLogFile

    try {
        if ($PSVersionTable.PSVersion.Major -ge 7) {

            Write-Log "Enumeration engine: .NET EnumerationOptions (PS7+)" -Level DEBUG -File $RuleLogFile
            $enumOptions = [System.IO.EnumerationOptions]::new()
            $enumOptions.RecurseSubdirectories = $true
            $enumOptions.IgnoreInaccessible    = $true
            $enumOptions.AttributesToSkip      = [System.IO.FileAttributes]::System -bor
                                                 [System.IO.FileAttributes]::ReparsePoint

            $rootDirInfo = [System.IO.DirectoryInfo]::new($root)
            foreach ($fi in $rootDirInfo.EnumerateFiles('*', $enumOptions)) {
                $result.FilesScanned++
                try {
                    $refDate = if ($Rule.UseCreationTime) { $fi.CreationTime } else { $fi.LastWriteTime }

                    if ($refDate -lt $cutoffDate -and (Test-ShouldInclude $fi $Rule)) {
                        $result.FilesMatched++

                        # Quota check before processing
                        if ($result.FilesDeleted -ge $Rule.MaxFiles) {
                            Write-Log "File quota reached ($($Rule.MaxFiles)). Stopping." -Level WARN -File $RuleLogFile
                            $result.QuotaReached = $true; break
                        }
                        if ($result.BytesDeleted + $fi.Length -gt $maxBytes) {
                            Write-Log "Volume quota reached ($($Rule.MaxDeleteMB) MB). Stopping." -Level WARN -File $RuleLogFile
                            $result.QuotaReached = $true; break
                        }

                        $age = [math]::Round(($StartTime - $refDate).TotalDays, 1)

                        if ($IsWhatIf) {
                            if ($result.FilesMatched -le 1000) {
                                Write-Log "[WHATIF] Would delete: $($fi.FullName)  (age: ${age}d, $(Format-Bytes $fi.Length))" -Level DEBUG -File $RuleLogFile
                            }
                            elseif ($result.FilesMatched -eq 1001) {
                                Write-Log "[WHATIF] ... (remaining matches not listed -- $(Format-Bytes ($maxBytes)) quota limit)" -Level DEBUG -File $RuleLogFile
                            }
                            $result.FilesDeleted++
                            $result.BytesDeleted += $fi.Length
                        }
                        else {
                            try {
                                Remove-Item -LiteralPath $fi.FullName -Force -ErrorAction Stop
                                Write-Log "Deleted: $($fi.FullName)  (age: ${age}d, $(Format-Bytes $fi.Length))" -Level SUCCESS -File $RuleLogFile
                                $result.FilesDeleted++
                                $result.BytesDeleted += $fi.Length
                                $result.ReportEntries.Add([PSCustomObject]@{
                                    Path = $fi.FullName; AgeDays = $age; SizeBytes = $fi.Length
                                    DeletedAt = (Get-Date -Format 'o'); Status = 'Deleted'
                                })
                            }
                            catch {
                                Write-Log "ERROR deleting: $($fi.FullName) -- $($_.Exception.Message)" -Level ERROR -File $RuleLogFile
                                $result.FilesErrored++
                                if ($result.ExitCode -lt 3) { $result.ExitCode = 3 }
                                $result.ReportEntries.Add([PSCustomObject]@{
                                    Path = $fi.FullName; AgeDays = $age; SizeBytes = $fi.Length
                                    DeletedAt = ''; Status = "Error: $($_.Exception.Message)"
                                })
                            }
                        }
                    }
                }
                catch { }

                if ($result.FilesScanned % 100000 -eq 0) {
                    Write-Log "$($result.FilesScanned.ToString('N0')) files scanned, $($result.FilesDeleted.ToString('N0')) deleted, $(Format-Bytes $result.BytesDeleted)..." -Level DEBUG -File $RuleLogFile
                }
            }
        }
        else {
            Write-Log "Enumeration engine: manual .NET Framework recursion + inline delete (PS5.1)" -Level DEBUG -File $RuleLogFile
            $script:ps51_Stats       = $result
            $script:ps51_Cutoff      = $cutoffDate
            $script:ps51_UseCreation = $Rule.UseCreationTime
            $script:ps51_Rule        = $Rule
            $script:ps51_LogFile     = $RuleLogFile
            $script:ps51_IsWhatIf    = $IsWhatIf
            $script:ps51_MaxBytes    = $maxBytes
            $script:ps51_StartTime   = $StartTime
            $script:ps51_Report      = $result.ReportEntries
            $script:ps51_DiagOldest  = [DateTime]::MaxValue
            $script:ps51_DiagNewest  = [DateTime]::MinValue
            $script:ps51_DiagDone    = $false

            Invoke-RecurseEnum51 -DirPath $root

            if ($result.FilesScanned -gt 0) {
                $dateF     = if ($Rule.UseCreationTime) { 'CreationTime' } else { 'LastWriteTime' }
                $oldestStr = if ($script:ps51_DiagOldest -eq [DateTime]::MaxValue) { 'N/A' } else { $script:ps51_DiagOldest.ToString('yyyy-MM-dd') }
                $newestStr = if ($script:ps51_DiagNewest -eq [DateTime]::MinValue) { 'N/A' } else { $script:ps51_DiagNewest.ToString('yyyy-MM-dd') }
                Write-Log "Diagnostic $dateF -- oldest: $oldestStr  |  newest: $newestStr  |  cutoff: $($cutoffDate.ToString('yyyy-MM-dd'))" -Level DEBUG -File $RuleLogFile
            }
        }
    }
    catch [UnauthorizedAccessException] {
        Write-Log "Access denied to root: $root" -Level ERROR -File $RuleLogFile
        $result.ExitCode = 1; return $result
    }
    catch {
        Write-Log "Enumeration error on $root : $($_.Exception.Message)" -Level ERROR -File $RuleLogFile
        $result.ExitCode = 1; return $result
    }

    Write-Log "Scan complete: $($result.FilesScanned.ToString('N0')) files scanned, $($result.FilesMatched.ToString('N0')) candidates, $($result.FilesDeleted.ToString('N0')) deleted, $(Format-Bytes $result.BytesDeleted)." -File $RuleLogFile

    if ($IsWhatIf -and $result.FilesMatched -gt 1000) {
        Write-Log "[WHATIF] Summary: $($result.FilesMatched.ToString('N0')) files would be deleted ($(Format-Bytes $result.BytesDeleted)) -- only first 1000 listed above." -Level WARN -File $RuleLogFile
    }

    # -- Empty folder cleanup -------------------------------------------------
    if ($Rule.PurgeEmptyFolders -and -not $result.QuotaReached) {
        Write-Log "Scanning for empty folders under: $root" -File $RuleLogFile
        $emptyFolders = Get-ChildItem -LiteralPath $root -Recurse -Directory -ErrorAction SilentlyContinue |
            Where-Object { (Get-ChildItem -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0 } |
            Sort-Object { $_.FullName.Length } -Descending
        foreach ($folder in $emptyFolders) {
            if ($IsWhatIf) {
                Write-Log "[WHATIF] Would remove empty folder: $($folder.FullName)" -Level DEBUG -File $RuleLogFile
                $result.FoldersDeleted++
            }
            else {
                try {
                    Remove-Item -LiteralPath $folder.FullName -Force -ErrorAction Stop
                    Write-Log "Empty folder removed: $($folder.FullName)" -Level SUCCESS -File $RuleLogFile
                    $result.FoldersDeleted++
                }
                catch { Write-Log "ERROR removing folder: $($folder.FullName) -- $($_.Exception.Message)" -Level WARN -File $RuleLogFile }
            }
        }
    }

    return $result
}

function Invoke-LogRotation {
    Write-Log "Rotating logs (retention: $LogRetentionDays days)..."
    $limit = $Script:StartTime.AddDays(-$LogRetentionDays)
    Get-ChildItem -Path $LogPath -File -Filter 'FilePurge_*' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -lt $limit } |
        ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force; Write-Log "Old log removed: $($_.Name)" -Level DEBUG }
            catch { Write-Log "Could not remove log: $($_.Name)" -Level WARN }
        }
}

function Write-EventLogEntry {
    param ([string] $Message, [string] $EntryType = 'Information', [int] $EventId = 1000)
    # Write-EventLog is Windows PowerShell 5.1 only.
    # On PS7+, use New-WinEvent if available, otherwise skip gracefully.
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        try {
            $logParams = @{
                ProviderName = $EventSource
                Id           = $EventId
                Payload      = @($Message)
            }
            New-WinEvent @logParams -ErrorAction Stop
        }
        catch {
            # Provider not registered or New-WinEvent unavailable -- skip silently
            Write-Log "Windows Event Log skipped (PS7): $($_.Exception.Message)" -Level DEBUG
        }
        return
    }
    # PS5.1 path
    try { New-EventLog -LogName Application -Source $EventSource -ErrorAction SilentlyContinue } catch { }
    try {
        Write-EventLog -LogName Application -Source $EventSource -EntryType $EntryType `
            -EventId $EventId -Message $Message
    }
    catch { Write-Log "Could not write to Windows Event Log: $($_.Exception.Message)" -Level WARN }
}

#endregion

# =============================================================================
# REGION : Main execution
# =============================================================================
#region Main

try {

    # -- Input validation -----------------------------------------------------
    $isJsonMode = -not [string]::IsNullOrWhiteSpace($ConfigFile)
    if (-not $isJsonMode -and ($null -eq $TargetPath -or $TargetPath.Count -eq 0)) {
        Write-Error "Either -ConfigFile or -TargetPath is required."
        exit 1
    }

    $effectiveWriteEventLog = [bool] $WriteEventLog
    if (-not $effectiveWriteEventLog -and $null -ne $script:jsonGlobal) {
        $effectiveWriteEventLog = [bool] (Get-JsonProp $script:jsonGlobal 'WriteEventLog' $false)
    }
    if (-not $PSBoundParameters.ContainsKey('LogRetentionDays') -and $null -ne $script:jsonGlobal) {
        $jLRD = Get-JsonProp $script:jsonGlobal 'LogRetentionDays' $null
        if ($null -ne $jLRD) { $LogRetentionDays = [int] $jLRD }
    }

    $useParallel = $Parallel -and ($PSVersionTable.PSVersion.Major -ge 7)
    if ($Parallel -and -not $useParallel) {
        Write-Log "-Parallel requires PS7+. Running sequentially on PS $($PSVersionTable.PSVersion)." -Level WARN
    }
    # -- Header ---------------------------------------------------------------
    Write-Log ('=' * 70) -Level SECTION
    Write-Log "INVOKE-FILEPURGE v$Script:Version  --  $(if ($Script:IsWhatIf) { 'SIMULATION MODE (WhatIf)' } else { 'REAL MODE' })" -Level SECTION
    Write-Log ('=' * 70) -Level SECTION
    Write-Log "Started       : $($Script:StartTime.ToString('yyyy-MM-dd HH:mm:ss'))"
    Write-Log "Mode          : $(if ($isJsonMode) { "JSON -- $ConfigFile" } else { 'CLI parameters' })"
    Write-Log "Execution     : $(if ($useParallel) { "Parallel (ThrottleLimit: $ThrottleLimit)" } else { 'Sequential' })"
    Write-Log "Log file      : $LogFile"
    Write-Log "Retention     : $LogRetentionDays days"

    # -- Load rules -----------------------------------------------------------
    [object[]] $rules = @(Resolve-PurgeRules)
    Write-Log "Rules loaded  : $($rules.Count) path(s) to process"

    # Re-evaluate parallel now that we know the rule count
    if ($useParallel -and $rules.Count -le 1) {
        $useParallel = $false
        Write-Log "-Parallel ignored: only $($rules.Count) rule(s) loaded -- running sequentially." -Level WARN
    }

    # -- Capture function bodies as STRINGS for parallel injection -------------
    # ForEach-Object -Parallel cannot receive ScriptBlock via $using: (PS7 restriction).
    # Workaround: serialize function bodies to strings, reconstruct inside each runspace.
    # Each rule writes to its own temp log -- no shared file, no mutex needed.
    if ($useParallel) {
        $fn_defs_str = (@(
            "function Safe-FilterArray { $( ${function:Safe-FilterArray} ) }"
            "function Normalize-Extensions { $( ${function:Normalize-Extensions} ) }"
            "function Test-RuleArray { $( ${function:Test-RuleArray} ) }"
            "function Get-JsonProp { $( ${function:Get-JsonProp} ) }"
            "function Get-FileAge { $( ${function:Get-FileAge} ) }"
            "function Format-Bytes { $( ${function:Format-Bytes} ) }"
            "function Write-Log { $( ${function:Write-Log} ) }"
            "function Test-ShouldInclude { $( ${function:Test-ShouldInclude} ) }"
            "function Invoke-RecurseEnum51 { $( ${function:Invoke-RecurseEnum51} ) }"
            "function Invoke-PurgeRule { $( ${function:Invoke-PurgeRule} ) }"
        ) -join "`n")
        $capturedIsWhatIf     = $Script:IsWhatIf
        $capturedOldestFirst = [bool]$OldestFirst
        $capturedStartTime = $Script:StartTime
        $capturedLogPath   = $LogPath
    }

    # -- Execute rules (sequential or parallel) -------------------------------
    $allResults = if ($useParallel) {

        $rules | ForEach-Object -ThrottleLimit $ThrottleLimit -Parallel {
            # Reconstruct all functions inside this isolated runspace
            . ([scriptblock]::Create($using:fn_defs_str))

            # Each rule writes to its own temp log segment (no shared file, no mutex)
            $tempLog = [System.IO.Path]::Combine($using:capturedLogPath,
                "FilePurge_rule_$([System.Guid]::NewGuid().ToString('N')).tmp")

            $Script:LogMutex  = [System.Threading.Mutex]::new($false)
            $Script:StartTime = $using:capturedStartTime

            $res = Invoke-PurgeRule `
                -Rule        $_ `
                -RuleLogFile $tempLog `
                -IsWhatIf    $using:capturedIsWhatIf `
                -StartTime   $using:capturedStartTime `
                -OldestFirst $using:capturedOldestFirst

            $Script:LogMutex.Dispose()

            # Tag result with temp log path for merging
            $res | Add-Member -NotePropertyName TempLogFile -NotePropertyValue $tempLog -Force
            $res
        }

    } else {

        foreach ($rule in $rules) {
            Invoke-PurgeRule `
                -Rule        $rule `
                -RuleLogFile $LogFile `
                -IsWhatIf    $Script:IsWhatIf `
                -StartTime   $Script:StartTime `
                -OldestFirst ([bool]$OldestFirst)
        }
    }

    # -- Merge results + temp log segments (parallel mode) -------------------
    foreach ($res in $allResults) {
        $GlobalStats.FilesScanned   += $res.FilesScanned
        $GlobalStats.FilesMatched   += $res.FilesMatched
        $GlobalStats.FilesDeleted   += $res.FilesDeleted
        $GlobalStats.FilesErrored   += $res.FilesErrored
        $GlobalStats.BytesDeleted   += $res.BytesDeleted
        $GlobalStats.FoldersDeleted += $res.FoldersDeleted
        if ($res.QuotaReached) { $GlobalStats.QuotaReached = $true }
        if ($res.ExitCode -gt $Script:ExitCode) { $Script:ExitCode = $res.ExitCode }
        foreach ($entry in $res.ReportEntries) { $GlobalReport.Add($entry) }

        # Merge per-rule temp log into main log (parallel mode only)
        $tmpLog = $null
        $tmpMember = $res | Get-Member -Name 'TempLogFile' -MemberType NoteProperty -ErrorAction SilentlyContinue
        if ($tmpMember) { $tmpLog = $res.TempLogFile }
        if (-not [string]::IsNullOrWhiteSpace($tmpLog) -and (Test-Path $tmpLog)) {
            Get-Content -Path $tmpLog -Encoding UTF8 | Add-Content -Path $LogFile -Encoding UTF8
            Remove-Item -Path $tmpLog -Force -ErrorAction SilentlyContinue
        }
    }

    # -- CSV report -----------------------------------------------------------
    if ($GlobalReport.Count -gt 0 -and -not $Script:IsWhatIf) {
        $GlobalReport | Export-Csv -Path $ReportFile -NoTypeInformation -Encoding UTF8
        Write-Log "CSV report exported: $ReportFile"
    }

    # -- Log rotation ---------------------------------------------------------
    Invoke-LogRotation

    # -- Global summary -------------------------------------------------------
    $Duration = New-TimeSpan -Start $Script:StartTime -End (Get-Date)
    $Summary  = @"
EXECUTION SUMMARY
  Mode           : $(if ($Script:IsWhatIf) { 'SIMULATION' } else { 'REAL' })
  Config         : $(if ($isJsonMode) { $ConfigFile } else { 'CLI parameters' })
  Execution      : $(if ($useParallel) { "Parallel (ThrottleLimit: $ThrottleLimit)" } else { 'Sequential' })
  Rules          : $($rules.Count) path(s)
  Duration       : $("{0:hh\:mm\:ss}" -f $Duration)
  Files scanned  : $($GlobalStats.FilesScanned.ToString('N0'))
  Candidates     : $($GlobalStats.FilesMatched.ToString('N0'))
  Deleted        : $($GlobalStats.FilesDeleted.ToString('N0'))
  Volume deleted : $(Format-Bytes $GlobalStats.BytesDeleted)
  Errors         : $($GlobalStats.FilesErrored)
  Empty folders  : $($GlobalStats.FoldersDeleted)
  Quota reached  : $(if ($GlobalStats.QuotaReached) { 'YES (partial purge)' } else { 'No' })
  Exit code      : $Script:ExitCode
"@

    Write-Log ('=' * 70) -Level SECTION
    $Summary.Split("`n") | ForEach-Object {
        Write-Log $_ -Level $(if ($_ -match 'Error|YES|SIMULATION') { 'WARN' } else { 'INFO' })
    }
    Write-Log ('=' * 70) -Level SECTION

    if ($effectiveWriteEventLog) {
        $entryType = switch ($Script:ExitCode) { 0 { 'Information' } 1 { 'Error' } default { 'Warning' } }
        Write-EventLogEntry -Message $Summary -EntryType $entryType -EventId (1000 + $Script:ExitCode)
    }

}
catch {
    Write-Log "UNHANDLED CRITICAL ERROR: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    $Script:ExitCode = 1
    if ($effectiveWriteEventLog) {
        Write-EventLogEntry -Message "CRITICAL ERROR FilePurge: $($_.Exception.Message)" -EntryType Error -EventId 1099
    }
}
finally {
    # Write-Log uses the mutex -- MUST be called before Dispose()
    Write-Log "Full log: $LogFile"
    try { $Script:LogMutex.Dispose() } catch { }
    exit $Script:ExitCode
}

#endregion
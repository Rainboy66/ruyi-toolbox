param([Parameter(Mandatory=$true)][string]$JobPath)
$ErrorActionPreference='Stop'
$job=Get-Content -LiteralPath $JobPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app=$null
$document=$null
$collection=$null
$owned=$false
$ownerRecords=@()
$ownerIds=@()
$stagePath=Join-Path ([IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($JobPath))) 'native-stage.json'
$kind=[string]$job.kind
if ($kind -notin @('word','sheet','slides')) { throw 'Unsupported WPS document kind.' }
$progid=@{word='KWPS.Application';sheet='KET.Application';slides='KWPP.Application'}[$kind]
if ($job.progid) { $progid=[string]$job.progid }

function Write-Stage([string]$stage,[string]$detail='') {
    @{stage=$stage;detail=$detail;kind=$kind;at=[DateTime]::UtcNow.ToString('o')} | ConvertTo-Json -Compress | Set-Content -LiteralPath $stagePath -Encoding UTF8
}
function Get-WpsProcesses {
    @(Get-CimInstance -ClassName Win32_Process -Filter "Name='wps.exe' OR Name='et.exe' OR Name='wpp.exe'")
}
function Is-AutomationProcess($process) {
    if (-not $process.CommandLine -or $process.CommandLine -notmatch '(?i)(?:/|-)(?:Automation|Embedding)(?:\s|$)') { return $false }
    $image=[string]$process.ExecutablePath
    if (-not $image) {
        try { $image=(Get-Process -Id $process.ProcessId -ErrorAction Stop).Path } catch { return $false }
    }
    return [string]::Equals([IO.Path]::GetFullPath($image),$serverPath,[StringComparison]::OrdinalIgnoreCase)
}
function Save-Ownership($processes) {
    $records=@()
    foreach ($process in $processes) {
        $native=Get-Process -Id $process.ProcessId -ErrorAction Stop
        if (-not [string]::Equals([IO.Path]::GetFullPath($native.Path),$serverPath,[StringComparison]::OrdinalIgnoreCase)) { throw 'WPS process identity changed before ownership was recorded.' }
        $created=$native.StartTime.ToUniversalTime().ToFileTimeUtc().ToString('D')
        $prior=@($ownerRecords | Where-Object { [int]$_.pid -eq $native.Id })
        if ($prior.Count -gt 0 -and $prior[0].created -ne $created) { throw 'WPS process identity changed after ownership was recorded.' }
        $records+=@{pid=$native.Id;created=$created}
    }
    if ($records.Count -eq 0) { throw 'No verified WPS process to record.' }
    $script:ownerRecords=$records
    $script:ownerIds=@($records | ForEach-Object { [int]$_.pid })
    @{pid=$records[0].pid;created=$records[0].created;owners=$records;engine='WPS Office';kind=$kind} | ConvertTo-Json -Compress -Depth 5 | Set-Content -LiteralPath $job.owner -Encoding UTF8
}
function Refresh-Ownership {
    if (-not $owned) { return }
    $current=@(Get-WpsProcesses)
    $fresh=@($current | Where-Object { $before -notcontains [int]$_.ProcessId -and (Is-AutomationProcess $_) })
    $accepted=@($fresh | Where-Object { $ownerIds -contains [int]$_.ProcessId })
    # Only descendants of the verified tree can be added after initial activation.
    do {
        $acceptedIds=@($accepted | ForEach-Object { [int]$_.ProcessId })
        $extra=@($fresh | Where-Object { $acceptedIds -notcontains [int]$_.ProcessId -and $acceptedIds -contains [int]$_.ParentProcessId })
        $accepted+=@($extra)
    } while ($extra.Count -gt 0)
    if ($accepted.Count -gt 0) { Save-Ownership $accepted }
}

try {
    Write-Stage 'activate'
    # The supervisor launches the PowerShell bitness matching WPS registration.
    $classes=[Microsoft.Win32.Registry]::ClassesRoot
    $key=$classes.OpenSubKey($progid+'\CLSID')
    if ($null -eq $key) { throw ('WPS automation class is not registered: '+$progid) }
    try { $clsid=[string]$key.GetValue('') } finally { $key.Close() }
    $key=$classes.OpenSubKey('CLSID\'+$clsid+'\LocalServer32')
    if ($null -eq $key) { throw 'WPS local automation server is not registered.' }
    try { $serverCommand=[string]$key.GetValue('') } finally { $key.Close() }
    $match=[regex]::Match($serverCommand,'^\s*"?(.+?\.exe)(?:"|\s|$)',[Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { throw 'Cannot resolve WPS automation executable.' }
    $serverPath=[IO.Path]::GetFullPath($match.Groups[1].Value)
    if (-not [IO.File]::Exists($serverPath)) { throw 'WPS automation executable is missing.' }
    $before=@(Get-WpsProcesses | ForEach-Object { [int]$_.ProcessId })
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class RuyiWpsWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
    $app=New-Object -ComObject $progid
    [uint32]$windowPid=0
    try { [void][RuyiWpsWindow]::GetWindowThreadProcessId([IntPtr]([long]$app.Hwnd),[ref]$windowPid) } catch {}
    switch ($kind) {
        'word' { $collection=$app.Documents }
        'sheet' { $collection=$app.Workbooks }
        'slides' { $collection=$app.Presentations }
    }
    if ($collection.Count -ne 0 -or ($windowPid -ne 0 -and $before -contains [int]$windowPid)) { throw 'WPS returned an existing document session; it was left untouched.' }
    $wait=[Diagnostics.Stopwatch]::StartNew()
    $stableIds=''
    $stablePasses=0
    $tree=@()
    do {
        $fresh=@(Get-WpsProcesses | Where-Object { $before -notcontains [int]$_.ProcessId -and (Is-AutomationProcess $_) })
        $freshIds=@($fresh | ForEach-Object { [int]$_.ProcessId })
        $roots=@($fresh | Where-Object { $freshIds -notcontains [int]$_.ParentProcessId })
        if ($roots.Count -eq 1 -and $before -notcontains [int]$roots[0].ParentProcessId) {
            $tree=@($roots[0])
            do {
                $treeIds=@($tree | ForEach-Object { [int]$_.ProcessId })
                $children=@($fresh | Where-Object { $treeIds -notcontains [int]$_.ProcessId -and $treeIds -contains [int]$_.ParentProcessId })
                $tree+=@($children)
            } while ($children.Count -gt 0)
            $treeIds=@($tree | ForEach-Object { [int]$_.ProcessId })
            $identity=(@($treeIds | Sort-Object) -join ',')
            # WPS spreadsheets hand the COM window to a sibling process created by
            # the same activation, so a tree-only check never matches and Excel
            # export always failed here. Accept a window whose process is new since
            # activation and runs the same executable; a user's old window still
            # fails ownership. Keep this file ASCII-only: 32-bit PowerShell reads
            # BOM-less scripts as ANSI and non-ASCII text breaks parsing.
            $windowOwned=($windowPid -eq 0 -or $treeIds -contains [int]$windowPid)
            if (-not $windowOwned) {
                $windowProcess=@(Get-WpsProcesses | Where-Object { [int]$_.ProcessId -eq [int]$windowPid })[0]
                if ($null -ne $windowProcess -and ($before -notcontains [int]$windowPid) -and [bool]$windowProcess.ExecutablePath) {
                    $windowExe=[IO.Path]::GetFullPath([string]$windowProcess.ExecutablePath)
                    $windowDir=[IO.Path]::GetDirectoryName($windowExe)
                    $serverDir=[IO.Path]::GetDirectoryName($serverPath)
                    $windowOwned=([string]::Equals($windowDir,$serverDir,[StringComparison]::OrdinalIgnoreCase)) -and
                        (@('wps.exe','et.exe','wpp.exe') -contains [IO.Path]::GetFileName($windowExe).ToLowerInvariant())
                }
            }
            if ($tree.Count -eq $fresh.Count -and $windowOwned) {
                if ($identity -eq $stableIds) { $stablePasses++ } else { $stablePasses=1;$stableIds=$identity }
                if ($stablePasses -ge 3) { break }
            } else { $stablePasses=0 }
        } else { $stablePasses=0 }
        Start-Sleep -Milliseconds 200
    } while ($wait.Elapsed.TotalSeconds -lt 8)
    if ($stablePasses -lt 3 -or $collection.Count -ne 0) { throw 'SESSION_CONFLICT: could not confirm one dedicated empty WPS session; existing sessions were left untouched.' }
    Save-Ownership $tree
    $owned=$true
    Write-Stage 'open'
    # Apply settings only after ownership; do not change fonts, spacing, print
    # areas, pagination, or persistent application options.
    $app.AutomationSecurity=3
    if ($kind -eq 'slides') { $app.DisplayAlerts=1 } else { $app.DisplayAlerts=0 }
    if ($kind -ne 'slides') { $app.Visible=$false }
    $sourcePath=[IO.Path]::GetFullPath($job.source)
    $outputPath=[IO.Path]::GetFullPath($job.output)
    switch ($kind) {
        'word' {
            $document=$collection.Open($sourcePath,$false,$true,$false)
            $document.Fields.Locked=$true
        }
        'sheet' {
            $document=$collection.Open($sourcePath,0,$true)
            # A raw data sheet often carries no print setup at all; exporting it
            # unchanged shrinks the whole range into a page corner and truncates
            # every cell to ###. Sheets that already define a print area keep the
            # author's settings; only unconfigured sheets get fit-to-width.
            foreach ($sheet in $document.Worksheets) {
                try {
                    if (-not $sheet.PageSetup.PrintArea) {
                        $sheet.PageSetup.Orientation = 2
                        $sheet.PageSetup.Zoom = $false
                        $sheet.PageSetup.FitToPagesWide = 1
                        $sheet.PageSetup.FitToPagesTall = $false
                    }
                } catch { }
            }
        }
        'slides' { $document=$collection.Open($sourcePath,-1,0,0) }
    }
    Refresh-Ownership
    Write-Stage 'export'
    switch ($kind) {
        'word' { $document.ExportAsFixedFormat($outputPath,17) }
        'sheet' { $document.ExportAsFixedFormat(0,$outputPath) }
        'slides' { $document.SaveAs($outputPath,32) }
    }
    Refresh-Ownership
    if (-not [IO.File]::Exists($outputPath)) { throw 'WPS returned without creating the PDF.' }
    @{engine='WPS Office';kind=$kind;progid=$progid} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.result -Encoding UTF8
    Write-Stage 'complete'
} catch {
    Write-Stage 'failed' $_.Exception.Message
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $document) {
        try { if ($kind -eq 'slides') { $document.Saved=-1;$document.Close() } else { $document.Close(0) } } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document) } catch {}
    }
    if ($null -ne $app) {
        if ($owned) {
            try { if ($collection.Count -eq 0) { if ($kind -eq 'word') { $app.Quit(0) } else { $app.Quit() } } } catch {}
        }
        if ($null -ne $collection) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($collection) } catch {} }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
    }
}

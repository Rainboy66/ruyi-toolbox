param([Parameter(Mandatory=$true)][string]$JobPath)
$ErrorActionPreference='Stop'
$job=Get-Content -LiteralPath $JobPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app=$null
$document=$null
$owned=$false
try {
    $before=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class RuyiWordWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
    $app=New-Object -ComObject Word.Application
    [uint32]$nativePid=0
    try { [void][RuyiWordWindow]::GetWindowThreadProcessId([IntPtr]([long]$app.Hwnd),[ref]$nativePid) } catch {}
    if ($nativePid -eq 0) {
        # Word can have no HWND before its first document. Confirm exactly
        # one new, empty process as in the established Word QA worker.
        $wait=[Diagnostics.Stopwatch]::StartNew()
        do {
            $new=@(Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id })
            if ($new.Count -eq 1) { $nativePid=[uint32]$new[0].Id; break }
            Start-Sleep -Milliseconds 100
        } while ($wait.Elapsed.TotalSeconds -lt 5)
    }
    if ($nativePid -eq 0 -or $before -contains $nativePid -or $app.Documents.Count -ne 0) { throw ('Could not confirm a dedicated empty Word session; existing sessions untouched. PID='+$nativePid+'; existing='+($before -contains $nativePid)+'; documents='+$app.Documents.Count) }
    $native=Get-Process -Id $nativePid
    $owned=$true
    @{pid=$native.Id;created=$native.StartTime.ToUniversalTime().ToFileTimeUtc().ToString('D')} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.owner -Encoding UTF8
    $app.Visible=$false
    $app.AutomationSecurity=3
    $app.DisplayAlerts=0
    # Operate only on the caller's preprocessed temporary DOCX. Do not change
    # persistent Word.Options or update fields/links from the user's document.
    $document=$app.Documents.Open([IO.Path]::GetFullPath($job.source),$false,$true,$false)
    $document.Fields.Locked=$true
    $document.ExportAsFixedFormat([IO.Path]::GetFullPath($job.output),17)
    @{engine='Microsoft Word'} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.result -Encoding UTF8
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $document) {
        try { $document.Close(0) } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $app) {
        if ($owned) { try { if ($app.Documents.Count -eq 0) { $app.Quit(0) } } catch {} }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

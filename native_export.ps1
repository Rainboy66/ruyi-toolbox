param([Parameter(Mandatory=$true)][string]$JobPath)
$ErrorActionPreference='Stop'
$job=Get-Content -LiteralPath $JobPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app=$null
$deck=$null
$owned=$false
try {
    if (Get-Process -Name POWERPNT -ErrorAction SilentlyContinue) { throw 'PowerPoint is already open; leaving that session untouched.' }
    $before=@(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue | ForEach-Object Id)
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class RuyiNativeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
    $app=New-Object -ComObject PowerPoint.Application
    [uint32]$nativePid=0
    [void][RuyiNativeWindow]::GetWindowThreadProcessId([IntPtr]$app.HWND,[ref]$nativePid)
    if ($nativePid -eq 0 -or $before -contains $nativePid -or $app.Presentations.Count -ne 0) { throw 'Could not obtain a dedicated empty PowerPoint session.' }
    $native=Get-Process -Id $nativePid
    $owned=$true
    @{pid=$nativePid;created=$native.StartTime.ToUniversalTime().ToFileTimeUtc()} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.owner -Encoding UTF8
    $app.AutomationSecurity=3
    $app.DisplayAlerts=1
    $deck=$app.Presentations.Open([IO.Path]::GetFullPath($job.source),-1,0,0)
    $deck.SaveAs([IO.Path]::GetFullPath($job.output),32)
    @{pages=$deck.Slides.Count;engine='Microsoft PowerPoint'} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.result -Encoding UTF8
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $deck) {
        try { $deck.Close() } catch {}
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($deck)
    }
    if ($null -ne $app) {
        if ($owned) { try { if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch {} }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
}

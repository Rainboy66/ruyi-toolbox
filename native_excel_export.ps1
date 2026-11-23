param([Parameter(Mandatory=$true)][string]$JobPath)
$ErrorActionPreference='Stop'
$job=Get-Content -LiteralPath $JobPath -Raw -Encoding UTF8 | ConvertFrom-Json
$app=$null
$books=$null
$workbook=$null
$owned=$false
try {
    $source=[IO.Path]::GetFullPath($job.source)
    $output=[IO.Path]::GetFullPath($job.output)
    if (-not [IO.File]::Exists($source)) { throw 'The source workbook does not exist.' }
    if ([String]::Equals($source,$output,[StringComparison]::OrdinalIgnoreCase)) { throw 'Source and PDF output must be different files.' }
    if ([IO.Path]::GetExtension($output) -ine '.pdf') { throw 'The output must be a PDF file.' }
    if ([IO.File]::Exists($output)) { throw 'The temporary PDF output already exists.' }

    $before=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public class RuyiExcelWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid); }'
    $app=New-Object -ComObject Excel.Application
    [uint32]$nativePid=0
    [void][RuyiExcelWindow]::GetWindowThreadProcessId([IntPtr]([long]$app.Hwnd),[ref]$nativePid)
    $books=$app.Workbooks
    if ($nativePid -eq 0 -or $before -contains $nativePid -or $books.Count -ne 0) { throw 'Could not confirm a dedicated empty Excel session; existing sessions untouched.' }
    $native=Get-Process -Id $nativePid
    if ($native.ProcessName -ine 'EXCEL') { throw 'The Excel window does not belong to an Excel process.' }
    $created=$native.StartTime.ToUniversalTime().ToFileTimeUtc().ToString('D')
    $owned=$true
    @{pid=$native.Id;created=$created} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.owner -Encoding UTF8

    $app.Visible=$false
    $app.AutomationSecurity=3
    $app.EnableEvents=$false
    $app.DisplayAlerts=$false
    $missing=[Type]::Missing
    # Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
    # IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify,
    # Converter, AddToMru. No links, macros, events or explicit refresh calls.
    # Callers must screen automatic external-data connections before opening.
    $workbook=$books.Open($source,0,$true,$missing,'','',$true,$missing,$missing,$false,$false,$missing,$false)
    if (-not $workbook.ReadOnly) { throw 'Excel did not open the workbook read-only.' }
    # Preserve the workbook's own print areas, page setup, scale and page breaks.
    # Do not force every sheet onto one page or recompute its print layout.
    $workbook.ExportAsFixedFormat(0,$output,0,$true,$false,$missing,$missing,$false)
    @{engine='Microsoft Excel'} | ConvertTo-Json -Compress | Set-Content -LiteralPath $job.result -Encoding UTF8
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) } catch {}
    }
    if ($null -ne $app) {
        if ($owned) { try { if ($null -ne $books -and $books.Count -eq 0) { $app.Quit() } } catch {} }
        if ($null -ne $books) { try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($books) } catch {} }
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app) } catch {}
    }
}

param(
    [Parameter(Mandatory = $true)][string]$Source,
    [Parameter(Mandatory = $true)][string]$Output
)

$ErrorActionPreference = "Stop"
$programIds = @("Excel.Application", "ket.Application", "et.Application")
$failures = @()

foreach ($programId in $programIds) {
    $application = $null
    $workbook = $null
    try {
        $application = New-Object -ComObject $programId
        $application.Visible = $false
        $application.DisplayAlerts = $false
        $workbook = $application.Workbooks.Open($Source)
        $workbook.SaveAs($Output, 51)
        $workbook.Close($false)
        $application.Quit()
        if (Test-Path -LiteralPath $Output) {
            exit 0
        }
    }
    catch {
        $failures += "${programId}: $($_.Exception.Message)"
    }
    finally {
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($workbook)
        }
        if ($null -ne $application) {
            try { $application.Quit() } catch {}
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($application)
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

[Console]::Error.WriteLine(($failures -join [Environment]::NewLine))
exit 1

#Requires -Version 7.6
<#
.SYNOPSIS
    Saves, plans, executes and resumes a shared conversion job.
.DESCRIPTION
    Creates TXT and HTML samples in a new task folder, saves the conversion
    settings as portable JSON, verifies that planning writes no PDF or state,
    and checks the completed and reused counts. The same JSON opens in Studio
    and runs with the officeimo CLI.
.PARAMETER OutputDirectory
    Parent directory for a new SavedConversionJob folder. That folder must not
    already exist so this example cannot overwrite an earlier run.
.EXAMPLE
    ./Recipe-SavedConversionJob.ps1 -OutputDirectory ./Output
.OUTPUTS
    The run summaries and the saved-job file.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
$root = Join-Path $OutputDirectory 'SavedConversionJob'
if (Test-Path -LiteralPath $root) {
    throw "The sample folder already exists: $root. Choose another output directory."
}
$source = Join-Path $root 'Source'
$pdf = Join-Path $root 'PDF'
$state = Join-Path $root 'State'
$job = Join-Path $root 'delivery.officejob.json'
$null = New-Item -ItemType Directory -Path $source -Force
'Plain text document for a repeatable delivery job.' | Set-Content -LiteralPath (Join-Path $source 'notes.txt')
'<!doctype html><html><body><h1>Delivery brief</h1><p>HTML document for the same job.</p></body></html>' |
    Set-Content -LiteralPath (Join-Path $source 'brief.html')

# Save the job instead of running it.
$null = Convert-OfficeDocument -Path $source -To Pdf -OutputDirectory $pdf -CheckpointPath $state -SaveJob $job

$plan = Convert-OfficeDocument -JobPath $job -Plan
if (-not $plan.CanRun -or $plan.Selected -ne 2) { throw 'The plan did not select both sample documents.' }
if ((Test-Path -LiteralPath $pdf) -or (Test-Path -LiteralPath $state)) { throw 'Planning created output or checkpoint state.' }

$firstReport = Join-Path $root 'first-run.json'
$null = Convert-OfficeDocument -JobPath $job -ReportPath $firstReport -SummaryVariable first
if ($first.Completed -ne 2 -or $first.Failed -ne 0) { throw 'The first run did not complete both sample documents.' }
$null = Get-OfficeDocumentInfo -Path (Join-Path $pdf '*.pdf')   # Throws if a PDF cannot be read back.
$recorded = Import-OfficeConversionReport -Path $firstReport
if ($recorded.Summary.Completed -ne 2 -or $recorded.Items.Count -ne 2) { throw 'The first attempt did not retain both file outcomes.' }
$firstHash = (Get-FileHash -LiteralPath $firstReport).Hash

# Running the job again reuses the verified outputs recorded in the checkpoint.
$null = Convert-OfficeDocument -JobPath $job -ReportPath (Join-Path $root 'resumed-run.json') -SummaryVariable resumed
if ($resumed.Completed -ne 2 -or $resumed.Reused -ne 2 -or $resumed.Failed -ne 0) { throw 'The second run did not reuse both verified outputs.' }
if ((Get-FileHash -LiteralPath $firstReport).Hash -ne $firstHash) { throw 'The second attempt changed the earlier run report.' }

$first
$resumed
Get-Item -LiteralPath $job

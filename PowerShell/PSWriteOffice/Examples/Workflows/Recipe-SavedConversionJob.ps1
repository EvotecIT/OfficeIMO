#Requires -Version 7.4
<#
.SYNOPSIS
    Saves, plans, executes and resumes a shared conversion job.
.DESCRIPTION
    Creates TXT and HTML samples in a new task folder, saves the conversion
    settings as portable JSON, verifies that planning writes no PDF/state, and
    checks the completed and reused counts. The same JSON can be opened in
    Studio or passed to officeimo workflow job plan/run.
.PARAMETER OutputDirectory
    Parent directory for a new SavedConversionJob folder. That folder must not
    already exist so this example cannot overwrite an earlier run.
.EXAMPLE
    ./Recipe-SavedConversionJob.ps1 -OutputDirectory ./Output
.OUTPUTS
    Conversion summary records and the saved-job file location.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $OutputDirectory
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop

$parent = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$root = Join-Path $parent 'SavedConversionJob'
if (Test-Path -LiteralPath $root) {
    throw "The sample folder already exists: $root. Choose another output directory."
}
$source = Join-Path $root 'Source'
$pdf = Join-Path $root 'PDF'
$state = Join-Path $root 'State'
$job = Join-Path $root 'delivery.officejob.json'
New-Item -ItemType Directory -Path $source -Force | Out-Null
'Plain text document for a repeatable delivery job.' |
    Set-Content -LiteralPath (Join-Path $source 'notes.txt') -Encoding utf8
'<!doctype html><html><body><h1>Delivery brief</h1><p>HTML document for the same job.</p></body></html>' |
    Set-Content -LiteralPath (Join-Path $source 'brief.html') -Encoding utf8

Export-OfficeConversionJob -InputDirectory $source -OutputDirectory $pdf `
    -CheckpointDirectory $state -TargetExtension pdf -Path $job | Out-Null

$plan = Invoke-OfficeConversionJob -Path $job -Plan
if (-not $plan.CanRun -or $plan.Selected -ne 2) {
    throw 'The plan did not select both sample documents.'
}
if ((Test-Path -LiteralPath $pdf) -or (Test-Path -LiteralPath $state)) {
    throw 'Planning created output or checkpoint state.'
}

$firstReportPath = Join-Path $root 'first-run.json'
$first = Invoke-OfficeConversionJob -Path $job -ReportPath $firstReportPath
if ($first.Completed -ne 2 -or $first.Failed -ne 0) {
    throw 'The first run did not complete both sample documents.'
}
foreach ($name in 'notes.txt.pdf', 'brief.html.pdf') {
    $path = Join-Path $pdf $name
    $document = Open-OfficePdf -Path $path
    if (-not $document) {
        throw "The completed PDF could not be reopened: $path"
    }
}
$recorded = Import-OfficeConversionReport -LiteralPath $firstReportPath
if ($recorded.Summary.Completed -ne 2 -or $recorded.Items.Count -ne 2) {
    throw 'The first attempt did not retain both file outcomes.'
}
$firstReportHash = (Get-FileHash -LiteralPath $firstReportPath -Algorithm SHA256).Hash

$resumed = Invoke-OfficeConversionJob -Path $job -ReportPath (Join-Path $root 'resumed-run.json')
if ($resumed.Completed -ne 2 -or $resumed.Reused -ne 2 -or $resumed.Failed -ne 0) {
    throw 'The second run did not reuse both verified artifacts.'
}
if ((Get-FileHash -LiteralPath $firstReportPath -Algorithm SHA256).Hash -ne $firstReportHash) {
    throw 'The second attempt changed the earlier run report.'
}
$first
$resumed
Get-Item -LiteralPath $job

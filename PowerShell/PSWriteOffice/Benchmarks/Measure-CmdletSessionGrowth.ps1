#Requires -Version 7.4
<#
.SYNOPSIS
Measures per-call cost of PSWriteOffice cmdlets that scripts call in loops, across a long session.

.DESCRIPTION
Each cmdlet runs in its own fresh PowerShell process against one live document, in batches of calls.
The script reports the per-call time of every batch and the gen 2 collections it caused, so a cost
that grows with session length (state left behind by earlier calls, a document operation that
rescans everything already added) shows up as a rising row. A trivial cmdlet gives the binding floor.

The module is loaded from this checkout's Release development binaries, so build them first:

    dotnet build Sources/PSWriteOffice/PSWriteOffice.csproj -c Release -f net8.0

.EXAMPLE
pwsh -NoLogo -NoProfile -File ./Benchmarks/Measure-CmdletSessionGrowth.ps1

.EXAMPLE
pwsh -NoLogo -NoProfile -File ./Benchmarks/Measure-CmdletSessionGrowth.ps1 -Cmdlet Add-OfficeMarkdownParagraph -Batches 4
#>
[CmdletBinding()]
param(
    [ValidateSet('New-OfficeMarkdownHtmlOptions', 'Convert-OfficeDocument', 'Add-OfficeMarkdownParagraph', 'Add-OfficeWordParagraph',
        'Set-OfficeExcelCell', 'Get-OfficeExcelTable', 'Get-OfficeExcelNamedRange', 'Split-OfficeDocument')]
    [string[]] $Cmdlet = @('New-OfficeMarkdownHtmlOptions', 'Convert-OfficeDocument', 'Add-OfficeMarkdownParagraph', 'Add-OfficeWordParagraph',
        'Set-OfficeExcelCell', 'Get-OfficeExcelTable', 'Get-OfficeExcelNamedRange', 'Split-OfficeDocument'),
    [ValidateRange(1, 100)][int] $Batches = 8,
    [ValidateRange(1, 100000)][int] $BatchSize = 1000,
    [switch] $AsObject
)

$ErrorActionPreference = 'Stop'
$moduleManifest = Join-Path (Split-Path $PSScriptRoot -Parent) 'PSWriteOffice.psd1'
$pwsh = (Get-Process -Id $PID).Path

# Runs in the child process: one cmdlet, one live document, batches of calls.
$probe = @'
param([string] $Manifest, [string] $Cmdlet, [int] $Batches, [int] $BatchSize)
$ErrorActionPreference = 'Stop'
$env:PSWRITEOFFICE_DEVELOPMENT_CONFIGURATION = 'Release'
Import-Module $Manifest
$work = Join-Path ([IO.Path]::GetTempPath()) ("officeimo-cmdlet-growth-" + $PID)
New-Item -ItemType Directory -Path $work -Force | Out-Null
try {
    switch ($Cmdlet) {
        'New-OfficeMarkdownHtmlOptions' { $operation = { New-OfficeMarkdownHtmlOptions } }
        'Convert-OfficeDocument' { $operation = { Convert-OfficeDocument -Text 'Deployment finished in **staging**.' -From Markdown -To Html } }
        'Add-OfficeMarkdownParagraph' { $document = New-OfficeMarkdown; $operation = { Add-OfficeMarkdownParagraph -Document $document -Text 'Reviewed and approved' } }
        'Add-OfficeWordParagraph' { $document = New-OfficeWord; $operation = { Add-OfficeWordParagraph -Document $document -Text 'Reviewed and approved' } }
        'Set-OfficeExcelCell' {
            $document = New-OfficeExcel
            $sheet = Add-OfficeExcelSheet -Document $document -Name Data -PassThru
            $script:row = 0
            $operation = { $script:row++; Set-OfficeExcelCell -Document $document -Sheet $sheet -Row $script:row -Column 1 -Value $script:row }
        }
        'Get-OfficeExcelTable' {
            $document = New-OfficeExcel
            $sheet = Add-OfficeExcelSheet -Document $document -Name Data -PassThru
            Add-OfficeExcelTable -Document $document -Sheet $sheet -Name 'Sales' -InputObject @(1..5 | ForEach-Object { [pscustomobject]@{ Region = "R$_"; Value = $_ } })
            $operation = { Get-OfficeExcelTable -Document $document }
        }
        'Get-OfficeExcelNamedRange' {
            $document = New-OfficeExcel
            $sheet = Add-OfficeExcelSheet -Document $document -Name Data -PassThru
            Set-OfficeExcelCell -Document $document -Sheet $sheet -Address A1 -Value 1
            Set-OfficeExcelNamedRange -Document $document -Name 'Total' -Range 'Data!$A$1'
            $operation = { Get-OfficeExcelNamedRange -Document $document }
        }
        'Split-OfficeDocument' {
            $pdf = Join-Path $work 'two-pages.pdf'
            Convert-OfficeDocument -Text "# One`n`nFirst page.`n`n---`n`n# Two" -From Markdown -To Pdf -OutputPath $pdf | Out-Null
            $script:split = 0
            $operation = { $script:split++; Split-OfficeDocument -Path $pdf -OutputDirectory (Join-Path $work ('split' + ($script:split % 4))) -PageCount 1 }
        }
    }

    # The untimed first call also checks that the cmdlets that return something did return it.
    $first = & $operation
    $returnsOutput = 'Convert-OfficeDocument', 'Get-OfficeExcelTable', 'Get-OfficeExcelNamedRange', 'New-OfficeMarkdownHtmlOptions', 'Split-OfficeDocument'
    if ($Cmdlet -in $returnsOutput -and $null -eq $first) { throw "$Cmdlet returned no output." }

    foreach ($batch in 1..$Batches) {
        $gen2 = [GC]::CollectionCount(2)
        $watch = [Diagnostics.Stopwatch]::StartNew()
        for ($i = 0; $i -lt $BatchSize; $i++) { $null = & $operation }
        [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0}|{1}|{2}', $batch * $BatchSize,
            $watch.Elapsed.TotalMilliseconds * 1000 / $BatchSize, [GC]::CollectionCount(2) - $gen2)
    }
} finally {
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}
'@

$probePath = Join-Path ([IO.Path]::GetTempPath()) ('officeimo-cmdlet-growth-probe-{0}.ps1' -f $PID)
Set-Content -LiteralPath $probePath -Value $probe
try {
    $results = foreach ($name in $Cmdlet) {
        $output = @(& $pwsh -NoLogo -NoProfile -File $probePath -Manifest $moduleManifest -Cmdlet $name -Batches $Batches -BatchSize $BatchSize)
        if ($LASTEXITCODE -ne 0) { throw "The $name run failed: $($output -join [Environment]::NewLine)" }
        $rows = @($output | Where-Object { $_ -match '^\d+\|' })
        if ($rows.Count -ne $Batches) { throw "The $name run returned $($rows.Count) of $Batches batches: $($output -join [Environment]::NewLine)" }
        foreach ($row in $rows) {
            $parts = $row -split '\|'
            [pscustomobject]@{
                Cmdlet               = $name
                CallsSoFar           = [int]$parts[0]
                MicrosecondsPerCall  = [Math]::Round([double]::Parse($parts[1], [Globalization.CultureInfo]::InvariantCulture), 1)
                Gen2Collections      = [int]$parts[2]
            }
        }
    }
} finally {
    Remove-Item -LiteralPath $probePath -ErrorAction SilentlyContinue
}

if ($AsObject) { return $results }

"PowerShell $($PSVersionTable.PSVersion); microseconds per call for each batch of $BatchSize calls in one session"
$table = foreach ($group in $results | Group-Object Cmdlet) {
    $row = [ordered]@{ Cmdlet = $group.Name }
    foreach ($entry in $group.Group) { $row[[string]$entry.CallsSoFar] = $entry.MicrosecondsPerCall }
    $row['Gen2'] = [int]($group.Group | Measure-Object Gen2Collections -Sum).Sum
    [pscustomobject]$row
}
$table | Format-Table -AutoSize | Out-String -Width 220

[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $InputDirectory,
    [Parameter(Mandatory)][string] $OutputDirectory,
    [ValidateSet('Word', 'PowerPoint', 'Excel')][string[]] $Applications = @('Word', 'PowerPoint', 'Excel')
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$inputRoot = (Resolve-Path -LiteralPath $InputDirectory).Path
$outputRoot = [IO.Path]::GetFullPath($OutputDirectory)
if (Test-Path -LiteralPath $outputRoot) { throw 'Choose a new output directory.' }
foreach ($name in @('report.docx', 'report.pptx', 'report.xlsx')) {
    if (-not (Test-Path -LiteralPath (Join-Path $inputRoot $name) -PathType Leaf)) { throw "Missing report: $name" }
}
if (@(Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue).Count -ne 0) {
    throw 'Office applications are already running. This isolated verification does not reuse or close existing application sessions.'
}
[void][IO.Directory]::CreateDirectory($outputRoot)
$results = [Collections.Generic.List[object]]::new()
$application = $document = $null
if ($Applications -contains 'Word') { try {
    Write-Output 'Word: starting application'
    $application = New-Object -ComObject Word.Application
    $application.Visible = $false
    $application.DisplayAlerts = 0
    $application.AutomationSecurity = 3
    Write-Output 'Word: opening report'
    $document = $application.Documents.Open((Join-Path $inputRoot 'report.docx'), $false, $true)
    Write-Output 'Word: exporting report'
    $document.Repaginate()
    $document.ExportAsFixedFormat((Join-Path $outputRoot 'word.pdf'), 17)
    $results.Add([pscustomobject]@{ Application = 'Word'; Version = [string]$application.Version; Tables = [int]$document.Tables.Count; Pages = [int]$document.ComputeStatistics(2); Text = [string]$document.Content.Text })
} finally {
    try { if ($null -ne $document) { try { $document.Close(0) } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null } } }
    finally { if ($null -ne $application) { try { $application.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application); $application = $null } } }
} }
if ($Applications -contains 'PowerPoint') { try {
    Write-Output 'PowerPoint: starting application'
    $application = New-Object -ComObject PowerPoint.Application
    $application.AutomationSecurity = 3
    Write-Output 'PowerPoint: opening report'
    $document = $application.Presentations.Open((Join-Path $inputRoot 'report.pptx'), -1, 0, 0)
    Write-Output 'PowerPoint: exporting report'
    $document.SaveAs((Join-Path $outputRoot 'powerpoint.pdf'), 32)
    $slides = [Collections.Generic.List[object]]::new()
    for ($index = 1; $index -le $document.Slides.Count; $index++) {
        $slide = $document.Slides.Item($index)
        try {
            $height = [int][Math]::Round(1440 * $document.PageSetup.SlideHeight / $document.PageSetup.SlideWidth)
            $slide.Export((Join-Path $outputRoot ("slide-{0}.png" -f $index)), 'PNG', 1440, $height)
            $slides.Add([pscustomobject]@{ Number = $index; Shapes = [int]$slide.Shapes.Count })
        } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($slide) }
    }
    $results.Add([pscustomobject]@{ Application = 'PowerPoint'; Version = [string]$application.Version; Slides = $slides.ToArray() })
} finally {
    try { if ($null -ne $document) { try { $document.Close() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null } } }
    finally { if ($null -ne $application) { try { $application.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application); $application = $null } } }
} }
if ($Applications -contains 'Excel') { try {
    Write-Output 'Excel: starting application'
    $application = New-Object -ComObject Excel.Application
    $application.Visible = $false
    $application.DisplayAlerts = $false
    $application.AutomationSecurity = 3
    Write-Output 'Excel: opening report'
    $document = $application.Workbooks.Open((Join-Path $inputRoot 'report.xlsx'), 0, $true)
    Write-Output 'Excel: exporting report'
    $document.ExportAsFixedFormat(0, (Join-Path $outputRoot 'excel.pdf'))
    $sheets = [Collections.Generic.List[object]]::new()
    for ($index = 1; $index -le $document.Worksheets.Count; $index++) {
        $sheet = $document.Worksheets.Item($index)
        try { $sheets.Add([pscustomobject]@{ Name = [string]$sheet.Name; Rows = [int]$sheet.UsedRange.Rows.Count; Columns = [int]$sheet.UsedRange.Columns.Count }) }
        finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    }
    $results.Add([pscustomobject]@{ Application = 'Excel'; Version = [string]$application.Version; Sheets = $sheets.ToArray() })
} finally {
    try { if ($null -ne $document) { try { $document.Close($false) } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document); $document = $null } } }
    finally { if ($null -ne $application) { try { $application.Quit() } finally { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application); $application = $null } } }
} }
$json = ConvertTo-Json -InputObject $results.ToArray() -Depth 6
[IO.File]::WriteAllText((Join-Path $outputRoot 'readback.json'), $json, [Text.UTF8Encoding]::new($false))
$json

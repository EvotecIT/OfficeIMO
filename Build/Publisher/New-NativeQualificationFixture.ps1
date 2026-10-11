[CmdletBinding()]
param([Parameter(Mandatory)][string] $OutputDirectory)

# Opt-in Windows oracle. Microsoft Publisher creates the input and reference PDF;
# neither OfficeIMO nor the normal build invokes this program.
$ErrorActionPreference = 'Stop'
$output = [IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Path $output -Force | Out-Null
$publication = Join-Path $output 'native-authored.pub'
$reference = Join-Path $output 'native-authored.pdf'
if ((Test-Path -LiteralPath $publication) -or (Test-Path -LiteralPath $reference)) {
    throw 'Choose an empty output directory; existing evidence is not overwritten.'
}
$application = $null
$document = $null
try {
    $application = New-Object -ComObject Publisher.Application
    $application.AutomationSecurity = 3
    $document = $application.NewDocument()
    $document.PageSetup.PageWidth = 612
    $document.PageSetup.PageHeight = 792
    $shapes = $document.Pages.Item(1).Shapes
    $text = $shapes.AddTextbox(1, 72, 36, 468, 36)
    $text.TextFrame.TextRange.Text = 'Native Publisher qualification'
    $text.TextFrame.TextRange.Font.Name = 'Arial'
    $text.TextFrame.TextRange.Font.Size = 18
    $gradient = $shapes.AddShape(1, 72, 96, 144, 72)
    $gradient.Fill.ForeColor.RGB = 255
    $gradient.Fill.BackColor.RGB = 16711680
    $gradient.Fill.TwoColorGradient(1, 1)
    $gradient.Line.Weight = 2
    $line = $shapes.AddLine(72, 216, 288, 216)
    $line.Line.Weight = 4
    $line.Line.BeginArrowheadStyle = 2
    $line.Line.EndArrowheadStyle = 5
    $line.Line.DashStyle = 4
    $diamond = $shapes.AddShape(4, 324, 96, 108, 72)
    $diamond.Rotation = 30
    $table = $shapes.AddTable(2, 2, 72, 288, 360, 100)
    $table.Table.Cells(1, 1, 1, 1).Item(1).TextRange.Text = 'First'
    $table.Table.Cells(1, 2, 1, 2).Item(1).TextRange.Text = 'Second'
    $table.Table.Cells(2, 1, 2, 1).Item(1).TextRange.Text = 'Alpha'
    $table.Table.Cells(2, 2, 2, 2).Item(1).TextRange.Text = 'Beta'
    $document.SaveAs($publication)
    $document.ExportAsFixedFormat(2, $reference)
    $version = [string] $application.Version
    $document.Close()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $null
    [ordered]@{
        producer = 'Microsoft Publisher'
        version = $version
        pages = 1
        widthPoints = 612
        heightPoints = 792
        publicationSha256 = (Get-FileHash -LiteralPath $publication -Algorithm SHA256).Hash.ToLowerInvariant()
        referenceSha256 = (Get-FileHash -LiteralPath $reference -Algorithm SHA256).Hash.ToLowerInvariant()
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output 'native-authored.json') -Encoding utf8
} finally {
    if ($null -ne $document) {
        $document.Close()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    }
    if ($null -ne $application) {
        $application.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($application)
    }
    foreach ($value in @($text, $gradient, $line, $diamond, $table, $shapes)) {
        if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

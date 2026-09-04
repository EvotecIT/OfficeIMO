[CmdletBinding()]
param(
    [string] $OutputDirectory,
    [ValidateSet('net8.0', 'net10.0')]
    [string] $Framework = 'net10.0',
    [ValidateSet('Debug', 'Release')]
    [string] $Configuration = 'Release'
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot 'Ignore\Benchmarks\PdfComparisons\office-com'
} elseif (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
    $OutputDirectory = Join-Path $repositoryRoot $OutputDirectory
}
$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
$fixtureDirectory = Join-Path $OutputDirectory 'files'
$projectPath = Join-Path $repositoryRoot 'OfficeIMO.Pdf.Benchmarks.Comparisons\OfficeIMO.Pdf.Benchmarks.Comparisons.csproj'
$docxPath = Join-Path $fixtureDirectory 'officeimo-word-rich-25-page.docx'
$wordPdfPath = Join-Path $fixtureDirectory 'microsoft-word-com-rich.pdf'
$xlsxPath = Join-Path $fixtureDirectory 'microsoft-excel-com-workbook.xlsx'
$excelPdfPath = Join-Path $fixtureDirectory 'microsoft-excel-com-workbook.pdf'
$pptxPath = Join-Path $fixtureDirectory 'microsoft-powerpoint-com-deck.pptx'
$powerPointPdfPath = Join-Path $fixtureDirectory 'microsoft-powerpoint-com-deck.pdf'
$additionalManifestPath = Join-Path $OutputDirectory 'office-com-corpus.json'

function Release-ComObject {
    param([AllowNull()][object] $Value)
    if ($null -ne $Value -and [System.Runtime.InteropServices.Marshal]::IsComObject($Value)) {
        try {
            [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)
        } catch {
            Write-Warning "Could not release COM object $($Value.GetType().FullName): $($_.Exception.Message)"
        }
    }
}

function Set-ExcelCellValue {
    param(
        [Parameter(Mandatory)][object] $Worksheet,
        [Parameter(Mandatory)][int] $Row,
        [Parameter(Mandatory)][int] $Column,
        [Parameter(Mandatory)][object] $Value
    )
    $cells = $null
    $cell = $null
    try {
        $cells = $Worksheet.Cells
        $cell = $cells.Item($Row, $Column)
        if ($Value -is [byte] -or
            $Value -is [short] -or
            $Value -is [int] -or
            $Value -is [long] -or
            $Value -is [single] -or
            $Value -is [double] -or
            $Value -is [decimal]) {
            $cell.Formula = '=' + [System.Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
        } else {
            $cell.Value2 = [string] $Value
        }
    } finally {
        Release-ComObject $cell
        Release-ComObject $cells
    }
}

function Set-PowerPointShapeText {
    param(
        [Parameter(Mandatory)][object] $Shape,
        [Parameter(Mandatory)][string] $Text,
        [single] $FontSize = 18
    )
    $textFrame = $null
    $textRange = $null
    $font = $null
    try {
        $textFrame = $Shape.TextFrame
        $textRange = $textFrame.TextRange
        $textRange.Text = $Text
        $font = $textRange.Font
        $font.Size = $FontSize
    } finally {
        Release-ComObject $font
        Release-ComObject $textRange
        Release-ComObject $textFrame
    }
}

function Add-PowerPointTextBox {
    param(
        [Parameter(Mandatory)][object] $Slide,
        [Parameter(Mandatory)][string] $Text,
        [Parameter(Mandatory)][single] $Left,
        [Parameter(Mandatory)][single] $Top,
        [Parameter(Mandatory)][single] $Width,
        [Parameter(Mandatory)][single] $Height,
        [single] $FontSize = 18
    )
    $shapes = $null
    $shape = $null
    try {
        $shapes = $Slide.Shapes
        $shape = $shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
        Set-PowerPointShapeText -Shape $shape -Text $Text -FontSize $FontSize
    } finally {
        Release-ComObject $shape
        Release-ComObject $shapes
    }
}

function Set-PowerPointTableCell {
    param(
        [Parameter(Mandatory)][object] $Table,
        [Parameter(Mandatory)][int] $Row,
        [Parameter(Mandatory)][int] $Column,
        [Parameter(Mandatory)][string] $Text
    )
    $cell = $null
    $shape = $null
    try {
        $cell = $Table.Cell($Row, $Column)
        $shape = $cell.Shape
        Set-PowerPointShapeText -Shape $shape -Text $Text -FontSize 13
    } finally {
        Release-ComObject $shape
        Release-ComObject $cell
    }
}

New-Item -ItemType Directory -Path $fixtureDirectory -Force | Out-Null

& dotnet run --project $projectPath -c $Configuration -f $Framework -- prepare-rich-word --repo-root $repositoryRoot --output $fixtureDirectory
if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $docxPath -PathType Leaf)) {
    throw "OfficeIMO.Word fixture generation failed: $docxPath"
}

$wordVersion = $null
$word = $null
$documents = $null
$document = $null
$anchorRange = $null
$smartArtLayouts = $null
$layout = $null
$wordShapes = $null
$shape = $null
$smartArt = $null
$nodes = $null
try {
    $word = New-Object -ComObject Word.Application
    $wordVersion = [string] $word.Version
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $documents = $word.Documents
    $document = $documents.Open($docxPath, $false, $false)

    $anchorRange = $document.Content
    $anchorRange.Collapse(0)
    $anchorRange.InsertBreak(7)
    $anchorRange.Collapse(0)
    $anchorRange.InsertAfter("COM SMARTART INTEROPERABILITY PAGE`r")
    $anchorRange.Collapse(0)

    $smartArtLayouts = $word.SmartArtLayouts
    $layout = $smartArtLayouts.Item(1)
    $wordShapes = $document.Shapes
    $shape = $wordShapes.AddSmartArt($layout, 50, 100, 450, 220, $anchorRange)
    $smartArt = $shape.SmartArt
    $nodes = $smartArt.AllNodes
    $labels = @('Collect', 'Validate', 'Publish')
    for ($index = 1; $index -le $labels.Count; $index++) {
        $node = $null
        $textFrame = $null
        $textRange = $null
        try {
            $node = $nodes.Item($index)
            $textFrame = $node.TextFrame2
            $textRange = $textFrame.TextRange
            $textRange.Text = $labels[$index - 1]
        } finally {
            Release-ComObject $textRange
            Release-ComObject $textFrame
            Release-ComObject $node
        }
    }

    $document.ExportAsFixedFormat($wordPdfPath, 17)
} finally {
    Release-ComObject $nodes
    Release-ComObject $smartArt
    Release-ComObject $shape
    Release-ComObject $wordShapes
    Release-ComObject $layout
    Release-ComObject $smartArtLayouts
    Release-ComObject $anchorRange
    try {
        if ($null -ne $document) {
            try {
                $document.Close(0)
            } finally {
                Release-ComObject $document
            }
        }
    } finally {
        Release-ComObject $documents
        if ($null -ne $word) {
            try {
                $word.Quit()
            } finally {
                Release-ComObject $word
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$excelVersion = $null
$excel = $null
$workbooks = $null
$workbook = $null
$worksheets = $null
$worksheet = $null
$usedRange = $null
$columns = $null
$sourceRange = $null
$chartObjects = $null
$chartObject = $null
$chart = $null
$chartTitle = $null
$pageSetup = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excelVersion = [string] $excel.Version
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbooks = $excel.Workbooks
    $workbook = $workbooks.Add()
    $worksheets = $workbook.Worksheets
    $worksheet = $worksheets.Item(1)
    $worksheet.Name = 'Operations'

    Set-ExcelCellValue $worksheet 1 1 'EXCEL COM INTEROPERABILITY'
    $headers = @('ACCOUNT', 'REGION', 'AMOUNT', 'STATUS')
    for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
        Set-ExcelCellValue $worksheet 3 ($columnIndex + 1) $headers[$columnIndex]
    }
    $rows = @(
        [pscustomobject]@{ Account = 'ACC-001-01'; Region = 'Sverige'; Amount = 1250.50; Status = 'Klar' },
        [pscustomobject]@{ Account = 'ACC-002-02'; Region = 'Polska'; Amount = 980.25; Status = 'Gotowe' },
        [pscustomobject]@{ Account = 'ACC-003-03'; Region = '日本'; Amount = 740.75; Status = '準備完了' },
        [pscustomobject]@{ Account = 'ACC-004-04'; Region = 'Россия'; Amount = 640.50; Status = 'Готово' }
    )
    for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
        $row = $rows[$rowIndex]
        Set-ExcelCellValue $worksheet ($rowIndex + 4) 1 ($row.Account)
        Set-ExcelCellValue $worksheet ($rowIndex + 4) 2 ($row.Region)
        Set-ExcelCellValue $worksheet ($rowIndex + 4) 3 ($row.Amount)
        Set-ExcelCellValue $worksheet ($rowIndex + 4) 4 ($row.Status)
    }

    $usedRange = $worksheet.UsedRange
    $columns = $usedRange.Columns
    [void] $columns.AutoFit()
    $sourceRange = $worksheet.Range('B3', 'C7')
    $chartObjects = $worksheet.ChartObjects()
    $chartObject = $chartObjects.Add(320, 20, 360, 220)
    $chart = $chartObject.Chart
    $chart.SetSourceData($sourceRange)
    $chart.HasTitle = $true
    $chartTitle = $chart.ChartTitle
    $chartTitle.Text = 'AMOUNT BY REGION'
    $pageSetup = $worksheet.PageSetup
    $pageSetup.Orientation = 2
    $pageSetup.Zoom = $false
    $pageSetup.FitToPagesWide = 1
    $pageSetup.FitToPagesTall = 1
    $pageSetup.PrintArea = '$A$1:$J$20'

    $workbook.SaveAs($xlsxPath, 51)
    $workbook.ExportAsFixedFormat(0, $excelPdfPath)
} finally {
    Release-ComObject $pageSetup
    Release-ComObject $chartTitle
    Release-ComObject $chart
    Release-ComObject $chartObject
    Release-ComObject $chartObjects
    Release-ComObject $sourceRange
    Release-ComObject $columns
    Release-ComObject $usedRange
    Release-ComObject $worksheet
    Release-ComObject $worksheets
    try {
        if ($null -ne $workbook) {
            try {
                $workbook.Close($false)
            } finally {
                Release-ComObject $workbook
            }
        }
    } finally {
        Release-ComObject $workbooks
        if ($null -ne $excel) {
            try {
                $excel.Quit()
            } finally {
                Release-ComObject $excel
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$powerPointVersion = $null
$powerPoint = $null
$presentations = $null
$presentation = $null
$slides = $null
try {
    $powerPoint = New-Object -ComObject PowerPoint.Application
    $powerPointVersion = [string] $powerPoint.Version
    $powerPoint.DisplayAlerts = 1
    $presentations = $powerPoint.Presentations
    $presentation = $presentations.Add()
    $slides = $presentation.Slides

    $slide1 = $null
    $slide1Shapes = $null
    $tableShape = $null
    $table = $null
    try {
        $slide1 = $slides.Add(1, 12)
        Add-PowerPointTextBox $slide1 'POWERPOINT COM INTEROPERABILITY' 36 20 650 45 26
        $slide1Shapes = $slide1.Shapes
        $tableShape = $slide1Shapes.AddTable(5, 4, 36, 90, 650, 230)
        $table = $tableShape.Table
        for ($columnIndex = 0; $columnIndex -lt $headers.Count; $columnIndex++) {
            Set-PowerPointTableCell $table 1 ($columnIndex + 1) $headers[$columnIndex]
        }
        for ($rowIndex = 0; $rowIndex -lt $rows.Count; $rowIndex++) {
            $row = $rows[$rowIndex]
            Set-PowerPointTableCell $table ($rowIndex + 2) 1 ($row.Account)
            Set-PowerPointTableCell $table ($rowIndex + 2) 2 ($row.Region)
            Set-PowerPointTableCell $table ($rowIndex + 2) 3 ([string] $row.Amount)
            Set-PowerPointTableCell $table ($rowIndex + 2) 4 ($row.Status)
        }
    } finally {
        Release-ComObject $table
        Release-ComObject $tableShape
        Release-ComObject $slide1Shapes
        Release-ComObject $slide1
    }

    $slide2 = $null
    try {
        $slide2 = $slides.Add(2, 12)
        Add-PowerPointTextBox $slide2 'MULTILINGUAL CONTENT' 36 20 650 45 26
        Add-PowerPointTextBox $slide2 "Svenska: Klar`rPolski: Gotowe`rРусский: Готово`r中文: 准备完毕`rالعربية: جاهز" 70 95 590 260 24
    } finally {
        Release-ComObject $slide2
    }

    $slide3 = $null
    $slide3Shapes = $null
    try {
        $slide3 = $slides.Add(3, 12)
        Add-PowerPointTextBox $slide3 'PROCESS DIAGRAM' 36 20 650 45 26
        $slide3Shapes = $slide3.Shapes
        $diagramLabels = @('Collect', 'Validate', 'Publish')
        for ($index = 0; $index -lt $diagramLabels.Count; $index++) {
            $diagramShape = $null
            try {
                $diagramShape = $slide3Shapes.AddShape(5, (60 + (210 * $index)), 150, 170, 90)
                Set-PowerPointShapeText $diagramShape ($diagramLabels[$index]) 22
            } finally {
                Release-ComObject $diagramShape
            }
        }
        Add-PowerPointTextBox $slide3 'Evidence-driven layout without language dictionaries' 95 285 530 55 18
    } finally {
        Release-ComObject $slide3Shapes
        Release-ComObject $slide3
    }

    $presentation.SaveAs($pptxPath, 24)
    $presentation.SaveAs($powerPointPdfPath, 32)
} finally {
    Release-ComObject $slides
    try {
        if ($null -ne $presentation) {
            try {
                $presentation.Close()
            } finally {
                Release-ComObject $presentation
            }
        }
    } finally {
        Release-ComObject $presentations
        if ($null -ne $powerPoint) {
            try {
                $powerPoint.Quit()
            } finally {
                Release-ComObject $powerPoint
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

$requiredFiles = @($wordPdfPath, $excelPdfPath, $powerPointPdfPath)
foreach ($requiredFile in $requiredFiles) {
    if (-not (Test-Path -LiteralPath $requiredFile -PathType Leaf)) {
        throw "Microsoft Office did not create the expected PDF: $requiredFile"
    }
}

$wordPageExpectations = @(
    for ($pageNumber = 1; $pageNumber -le 25; $pageNumber++) {
        $expectation = [ordered]@{
            pageNumber = $pageNumber
            expectedTables = 1
            expectedImages = if ($pageNumber -eq 3) { 1 } else { 0 }
            expectedImageRegions = if ($pageNumber -eq 3) { 1 } else { 0 }
            expectedFigures = if ($pageNumber -eq 3) { 1 } else { 0 }
            minimumVectorPrimitives = 1
            tables = @(
                [ordered]@{
                    rows = 6
                    columns = 4
                    requiredCells = @(
                        'Account',
                        ('ACC-{0:D3}-01' -f $pageNumber),
                        ('ACC-{0:D3}-05' -f $pageNumber)
                    )
                }
            )
        }
        $expectation
    }
)
$wordPageExpectations += [ordered]@{
    pageNumber = 26
    expectedTables = 0
    expectedImages = 0
    expectedImageRegions = 0
    expectedFigures = 0
    minimumVectorPrimitives = 1
}

$manifest = [ordered]@{
    schemaVersion = 1
    description = 'Machine-local Microsoft Office COM PDF exports generated for parser validation.'
    entries = @(
        [ordered]@{
            id = 'microsoft-word-com-rich'
            sourceKind = 'local'
            sourcePath = $wordPdfPath
            sha256 = (Get-FileHash -LiteralPath $wordPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
            producer = "Microsoft Word $wordVersion COM export"
            license = 'Generated local fixture'
            tier = 'large'
            expectedPages = 26
            minimumTokenRecall = 0.90
            features = @('tables', 'chart', 'native-word-smartart', 'image', 'links', 'headers-footers', 'office-com-export')
            requiredText = @('Quarterly delivery', 'COM SMARTART INTEROPERABILITY PAGE', 'Collect', 'Validate', 'Publish')
            pageExpectations = $wordPageExpectations
        },
        [ordered]@{
            id = 'microsoft-excel-com-workbook'
            sourceKind = 'local'
            sourcePath = $excelPdfPath
            sha256 = (Get-FileHash -LiteralPath $excelPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
            producer = "Microsoft Excel $excelVersion COM export"
            license = 'Generated local fixture'
            tier = 'medium'
            expectedPages = 1
            minimumTokenRecall = 0.95
            features = @('spreadsheet', 'table', 'chart', 'multilingual', 'office-com-export')
            requiredText = @('EXCEL COM INTEROPERABILITY', 'ACC-001-01', 'ACC-004-04')
            expectedText = @('Sverige', 'Polska', '日本', 'Россия')
            pageExpectations = @(
                [ordered]@{
                    pageNumber = 1
                    expectedTables = 1
                    expectedImages = 0
                    expectedImageRegions = 0
                    expectedFigures = 0
                    minimumVectorPrimitives = 1
                    tables = @(
                        [ordered]@{
                            rows = 5
                            columns = 4
                            requiredCells = @('ACCOUNT', 'ACC-001-01', 'ACC-004-04')
                        }
                    )
                }
            )
        },
        [ordered]@{
            id = 'microsoft-powerpoint-com-deck'
            sourceKind = 'local'
            sourcePath = $powerPointPdfPath
            sha256 = (Get-FileHash -LiteralPath $powerPointPdfPath -Algorithm SHA256).Hash.ToLowerInvariant()
            producer = "Microsoft PowerPoint $powerPointVersion COM export"
            license = 'Generated local fixture'
            tier = 'medium'
            expectedPages = 3
            minimumTokenRecall = 0.95
            features = @('slides', 'table', 'diagram', 'multilingual', 'office-com-export')
            requiredText = @('POWERPOINT COM INTEROPERABILITY', 'MULTILINGUAL CONTENT', 'PROCESS DIAGRAM', 'Collect', 'Validate', 'Publish')
            expectedText = @('Svenska: Klar', 'Polski: Gotowe', 'Русский: Готово', '中文: 准备完毕', 'العربية: جاهز')
            pageExpectations = @(
                [ordered]@{
                    pageNumber = 1
                    expectedTables = 1
                    expectedImages = 0
                    expectedImageRegions = 0
                    expectedFigures = 0
                    minimumVectorPrimitives = 1
                    tables = @(
                        [ordered]@{
                            rows = 5
                            columns = 4
                            requiredCells = @('ACCOUNT', 'ACC-001-01', 'ACC-004-04')
                        }
                    )
                },
                [ordered]@{
                    pageNumber = 2
                    expectedTables = 0
                    expectedImages = 0
                    expectedImageRegions = 0
                    expectedFigures = 0
                    minimumVectorPrimitives = 1
                },
                [ordered]@{
                    pageNumber = 3
                    expectedTables = 0
                    expectedImages = 0
                    expectedImageRegions = 0
                    expectedFigures = 0
                    minimumVectorPrimitives = 3
                }
            )
        }
    )
}
$manifestJson = $manifest | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText($additionalManifestPath, $manifestJson, [System.Text.UTF8Encoding]::new($false))

$validationDirectory = Join-Path $OutputDirectory 'validation'
& dotnet run --project $projectPath -c $Configuration -f $Framework -- corpus `
    --repo-root $repositoryRoot `
    --additional-manifest $additionalManifestPath `
    --only 'microsoft-word-com-rich,microsoft-excel-com-workbook,microsoft-powerpoint-com-deck' `
    --skip-manipulation `
    --output $validationDirectory
if ($LASTEXITCODE -ne 0) {
    throw "Microsoft Office COM PDF corpus validation failed with exit code $LASTEXITCODE."
}

Write-Host "Microsoft Office COM PDF corpus validation passed: $validationDirectory"

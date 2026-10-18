Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Example-ExcelAdvanced.xlsx'
$data = @(
    [pscustomobject]@{ Region = 'North'; Quarter = 'Q1'; Sales = 1200; Status = 'New' }
    [pscustomobject]@{ Region = 'North'; Quarter = 'Q2'; Sales = 900; Status = 'Done' }
    [pscustomobject]@{ Region = 'South'; Quarter = 'Q1'; Sales = 1500; Status = 'In Progress' }
    [pscustomobject]@{ Region = 'South'; Quarter = 'Q2'; Sales = 700; Status = 'New' }
    [pscustomobject]@{ Region = 'West'; Quarter = 'Q1'; Sales = 1800; Status = 'Done' }
    [pscustomobject]@{ Region = 'West'; Quarter = 'Q2'; Sales = 1100; Status = 'In Progress' }
)

$imagePath = Join-Path $PSScriptRoot '..\Word\Example-WordTableCells.fixture.png'

New-OfficeExcel -Path $path {
    ExcelSheet 'Data' {
        ExcelTable -Data $data -Start A1 -AutoFit -Style 'TableStyleMedium9'
        ExcelValidationList -Range 'D2:D20' -Values 'New','In Progress','Done'
        ExcelConditionalRule -Range 'C2:C20' -Operator GreaterThan -Value 1000 -Style 'Good'
        ExcelAutoFilter -Range 'A1:D20'
        ExcelFreeze -Rows 1
        ExcelSort -Range 'A1:D20' -Key 'C' -Descending
        ExcelSparkline -DataRange 'C2:C7' -LocationRange 'E2:E7' -Type Line -ShowMarkers
        ExcelPivotTable -SourceRange 'A1:D7' -DestinationCell 'G2' -RowField 'Region' -DataField 'Sales'
        ExcelChart -Range 'A1:C7' -Type Column -Title 'Sales by Region' -Row 1 -Column 9
        ExcelHyperlink -Cell 'A2' -Uri 'https://evotec.xyz'
        ExcelComment -Cell 'C2' -Text 'Review this value'

        if (Test-Path $imagePath) {
            ExcelImage -Path $imagePath -Range 'I8:J12' -Name 'OfficeIMOLogo' -AltText 'OfficeIMO logo'
        }
    }

    ExcelSheet 'Print' {
        ExcelTable -Data $data -Start A1 -AutoFit
        ExcelOrientation -Orientation Landscape
        ExcelMargins -Left 0.5 -Right 0.5 -Top 0.75 -Bottom 0.75
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
        ExcelGridlines -Show $false
    }

    ExcelSheet 'Hidden' {
        ExcelCell -Address A1 -Value 'Internal notes'
        ExcelSheetVisibility -Hide
    }
} -Open

Write-Host "Workbook saved to $path"

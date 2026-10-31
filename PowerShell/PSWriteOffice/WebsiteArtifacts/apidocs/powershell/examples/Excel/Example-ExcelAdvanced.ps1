param([switch] $Show)

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
        ExcelTable -Data $data -Name 'SalesData' -AutoFit -TableStyle 'TableStyleMedium9'
        ExcelValidationList -Range 'D2:D20' -Values 'New','In Progress','Done'
        ExcelAutoFilter -Range 'A1:D20'
        ExcelSheetView -FreezeRows 1
        ExcelSort -Header 'Sales' -Descending
        ExcelSparkline -DataRange 'C2:C7' -LocationRange 'E2:E7' -Type Line -ShowMarkers
        ExcelPivotTable -SourceRange 'A1:D7' -DestinationCell 'G2' -RowField 'Region' -DataField 'Sales'
        ExcelChart -Table 'SalesData' -Category Region -Value Sales -Type ColumnClustered -Title 'Sales by Region' -Address I1
        ExcelHyperlink -Address 'A10' -Url 'https://evotec.xyz' -Display 'Evotec'
        ExcelComment -Address 'C2' -Text 'Review this value'

        if (Test-Path $imagePath) {
            ExcelImage -Path $imagePath -Range 'I20:J24' -Name 'OfficeIMOLogo' -AltText 'OfficeIMO logo'
        }
    } -PassThru | ForEach-Object {
        $null = $_.AddConditionalRule('C2:C20', [OfficeIMO.Excel.ExcelConditionalFormattingOperator]::GreaterThan,
            '1000', $null, '#C6EFCE')
    }

    ExcelSheet 'Print' {
        ExcelTable -Data $data -AutoFit
        ExcelOrientation -Orientation Landscape
        ExcelMargins -Left 0.5 -Right 0.5 -Top 0.75 -Bottom 0.75
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
        ExcelSheetView -NoGridlines
    }

    ExcelSheet 'Hidden' {
        ExcelCell -Address A1 -Value 'Internal notes'
        Set-OfficeExcelSheet -Visibility Hidden
    }
} -Show:$Show

Write-Host "Workbook saved to $path"

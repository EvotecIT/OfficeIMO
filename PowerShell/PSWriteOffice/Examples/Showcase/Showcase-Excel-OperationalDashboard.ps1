param(
    [switch] $Show,
    [string] $OutputDirectory = (Join-Path $PSScriptRoot 'Output')
)

$ErrorActionPreference = 'Stop'
Import-Module PSWriteOffice -ErrorAction Stop

$documents = $OutputDirectory
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Showcase-Excel-OperationalDashboard.xlsx'
$logoPath = Join-Path $PSScriptRoot '..\Word\Example-WordTableCells.fixture.png'

$services = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'Platform'; HealthScore = 96; Incidents = 1; SlaPercent = 0.999; Status = 'Healthy'; Evidence = 'identity-runbook' }
    [pscustomobject]@{ Service = 'Mail Flow'; Owner = 'Messaging'; HealthScore = 89; Incidents = 4; SlaPercent = 0.995; Status = 'Watch'; Evidence = 'mail-flow-kpis' }
    [pscustomobject]@{ Service = 'Endpoint Fleet'; Owner = 'Workplace'; HealthScore = 82; Incidents = 7; SlaPercent = 0.991; Status = 'Watch'; Evidence = 'endpoint-drift' }
    [pscustomobject]@{ Service = 'Backup'; Owner = 'Operations'; HealthScore = 74; Incidents = 11; SlaPercent = 0.986; Status = 'Risk'; Evidence = 'backup-exceptions' }
    [pscustomobject]@{ Service = 'Network Edge'; Owner = 'Network'; HealthScore = 93; Incidents = 2; SlaPercent = 0.997; Status = 'Healthy'; Evidence = 'edge-availability' }
    [pscustomobject]@{ Service = 'Data Warehouse'; Owner = 'Data'; HealthScore = 78; Incidents = 8; SlaPercent = 0.988; Status = 'Risk'; Evidence = 'warehouse-latency' }
    [pscustomobject]@{ Service = 'Collaboration'; Owner = 'Productivity'; HealthScore = 91; Incidents = 3; SlaPercent = 0.996; Status = 'Healthy'; Evidence = 'collab-telemetry' }
    [pscustomobject]@{ Service = 'Security Signals'; Owner = 'Security'; HealthScore = 87; Incidents = 5; SlaPercent = 0.993; Status = 'Watch'; Evidence = 'signal-coverage' }
)

$trend = @(
    [pscustomobject]@{ Month = 'Jan'; Availability = 99.1; Incidents = 14; Automation = 62 }
    [pscustomobject]@{ Month = 'Feb'; Availability = 99.2; Incidents = 12; Automation = 66 }
    [pscustomobject]@{ Month = 'Mar'; Availability = 99.4; Incidents = 10; Automation = 71 }
    [pscustomobject]@{ Month = 'Apr'; Availability = 99.3; Incidents = 11; Automation = 74 }
    [pscustomobject]@{ Month = 'May'; Availability = 99.6; Incidents = 8; Automation = 79 }
    [pscustomobject]@{ Month = 'Jun'; Availability = 99.7; Incidents = 6; Automation = 83 }
)

$legend = @(
    [pscustomobject]@{ Status = 'Healthy'; Meaning = 'Stable service posture'; Action = 'Keep monitoring' }
    [pscustomobject]@{ Status = 'Watch'; Meaning = 'Needs routine follow-up'; Action = 'Review trend and owners' }
    [pscustomobject]@{ Status = 'Risk'; Meaning = 'Needs active remediation'; Action = 'Open mitigation plan' }
)

$statusMix = $services |
    Group-Object Status |
    ForEach-Object {
        [pscustomobject]@{
            Status = $_.Name
            Count  = $_.Count
        }
    }

$serviceRows = $services |
    Select-Object Service, HealthScore, Incidents, Owner, SlaPercent, Status, Evidence

$ownerSummary = $services |
    Group-Object Owner |
    ForEach-Object {
        [pscustomobject]@{
            Owner         = $_.Name
            Services      = $_.Count
            AverageHealth = [math]::Round(($_.Group | Measure-Object HealthScore -Average).Average, 1)
            Incidents     = ($_.Group | Measure-Object Incidents -Sum).Sum
            RiskServices  = @($_.Group | Where-Object Status -eq 'Risk').Count
        }
    }

New-OfficeExcel -Path $path {
    ExcelSheet 'Summary' {
        ExcelSheetView -NoGridlines
        ExcelCell -Address 'A1' -Value 'Operational Dashboard'
        ExcelCell -Address 'A2' -Value 'Generated with PSWriteOffice from PowerShell objects.'
        ExcelCell -Address 'A4' -Value 'Average Health'
        ExcelCell -Address 'B4' -Formula 'AVERAGE(Services!B2:B9)' -NumberFormat '0.0'
        ExcelCell -Address 'D4' -Value 'Total Incidents'
        ExcelCell -Address 'E4' -Formula 'SUM(Services!C2:C9)' -NumberFormat '0'
        ExcelCell -Address 'G4' -Value 'Risk Services'
        ExcelCell -Address 'H4' -Formula 'COUNTIF(Services!F2:F9,"Risk")' -NumberFormat '0'

        ExcelTable -InputObject $legend -Name 'StatusLegend' -Address A7 -TableStyle 'TableStyleMedium4'
        ExcelTable -InputObject $statusMix -Name 'StatusMix' -Address F7 -TableStyle 'TableStyleMedium4'
        ExcelColumn -Column 4 -AutoFit
        ExcelChart -Range 'F7:G10' -Address I7 -Type Doughnut -Title 'Status Mix' -Width 440px -Height 260px -PassThru |
            Set-OfficeExcelChartLegend -Position Right -PassThru |
            Set-OfficeExcelChartDataLabel -ShowValue -ShowCategoryName -Position OutsideEnd -PassThru |
            Set-OfficeExcelChart -Style 251 -ColorStyle 10

        if (Test-Path $logoPath) {
            ExcelImage -Path $logoPath -Address 'J1' -Width 140px -Height 52px -AltText 'PSWriteOffice operational dashboard logo'
        }

        ExcelHeader -Center 'PSWriteOffice operational dashboard'
        ExcelFooter -Right 'Page &P of &N'
        ExcelPageSetup -Orientation Landscape
        ExcelPageSetup -MarginLeft 0.4in -MarginRight 0.4in -MarginTop 0.7in -MarginBottom 0.7in
        ExcelPageSetup -FitToWidth 1 -FitToHeight 0
    }

    ExcelSheet 'Services' {
        ExcelTable -InputObject $serviceRows -Name 'ServiceHealth' -Address A1 -TableStyle 'TableStyleMedium9'
        ExcelSheetView -FreezeRows 1
        ExcelValidation -Type List -Range 'F2:F50' -Values 'Healthy','Watch','Risk'
        ExcelConditionalFormat -Range 'B2:B9' -StartColor '#F8696B' -EndColor '#63BE7B'
        ExcelConditionalFormat -DataBar -Range 'C2:C9' -Color '#5B9BD5'
        ExcelConditionalFormat -Range 'B2:B9' -IconSet ThreeTrafficLights1
        ExcelHyperlink -Header 'Evidence' -Table 'ServiceHealth' -UrlScript { param($text) "https://evotec.xyz/docs/$text" } -DisplayScript { param($text) "Open $text" }
        ExcelPivotTable -Range 'A1:F9' -Address 'J1' -Name 'ServiceStatusPivot' -RowField Status -DataField Incidents -DataDisplayName 'Total Incidents' -Style PivotStyleMedium9 -RefreshOnOpen
        ExcelChart -Range 'A1:B9' -Address A12 -Type BarClustered -Title 'Service health score (0-100)' -Width 760px -Height 340px -PassThru |
            Set-OfficeExcelChartLegend -Position Bottom -PassThru |
            Set-OfficeExcelChartDataLabel -ShowValue -Position OutsideEnd -PassThru |
            Set-OfficeExcelChart -Style 251 -ColorStyle 10
        ExcelHeader -Center 'Service details'
        ExcelFooter -Right 'Page &P of &N'
    }

    ExcelSheet 'Trend' {
        ExcelTable -InputObject $trend -Name 'TrendData' -Address A1 -TableStyle 'TableStyleMedium2'
        ExcelSheetView -FreezeRows 1
        ExcelCell -Address 'F1' -Value 'Six-month trends'
        ExcelCell -Address 'F2' -Value 'Availability (%)'
        ExcelSparkline -DataRange 'B2:B7' -LocationRange 'G2'
        ExcelCell -Address 'F3' -Value 'Incidents'
        ExcelSparkline -DataRange 'C2:C7' -LocationRange 'G3'
        ExcelCell -Address 'F4' -Value 'Automation (%)'
        ExcelSparkline -DataRange 'D2:D7' -LocationRange 'G4'
        ExcelTable -InputObject ($trend | Select-Object Month, Incidents) -Name 'IncidentTrend' -Address J1 -TableStyle 'TableStyleMedium2'
        ExcelChart -Range 'J1:K7' -Address A10 -Type Line -Title 'Recorded incidents by month' -Width 780px -Height 340px -PassThru |
            Set-OfficeExcelChartLegend -Position Bottom -PassThru |
            Set-OfficeExcelChartDataLabel -ShowValue -Position Top -PassThru |
            Set-OfficeExcelChart -Style 251 -ColorStyle 10
        ExcelHeader -Center 'Trend and automation'
        ExcelFooter -Right 'Page &P of &N'
    }

    ExcelSheet 'Owner Summary' {
        ExcelTable -InputObject $ownerSummary -Name 'OwnerSummary' -Address A1 -TableStyle 'TableStyleMedium5'
        ExcelConditionalFormat -DataBar -Range 'D2:D20' -Color '#ED7D31'
        ExcelConditionalFormat -Range 'C2:C20' -IconSet ThreeTrafficLights1
        ExcelHeader -Center 'Owner summary'
        ExcelFooter -Right 'Page &P of &N'
    }

    ExcelSheet 'Notes' {
        ExcelCell -Address 'A1' -Value 'Generation Notes'
        ExcelCell -Address 'A2' -Value 'This sheet is hidden and carries inputs useful for audit/debugging.'
        ExcelCell -Address 'A4' -Value 'Created'
        ExcelCell -Address 'B4' -Value (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        ExcelCell -Address 'A5' -Value 'Source'
        ExcelCell -Address 'B5' -Value 'Examples/Showcase/Showcase-Excel-OperationalDashboard.ps1'
        Set-OfficeExcelSheet -Visibility Hidden
    }

    ExcelTableOfContents -Name 'Index' -IncludeNamedRanges -BackLink -BackLinkText 'Back to Index'
} -Show:$Show -SaveOptions (New-OfficeExcelSaveOptions -EvaluateFormulas)

Open-OfficeExcel -Path $path {
    $threaded = Add-OfficeExcelComment -Sheet Summary -Address A2 -Text 'Review dashboard posture before sending to service owners.' -Author 'Automation Reviewer' -Threaded -PassThru
    Add-OfficeExcelComment -Sheet Summary -Address A2 -Text 'Ready for owner review.' -Author 'Report Owner' -Threaded -ParentId $threaded.Id -Done

    Add-OfficeExcelConnection -Name 'OperationalDashboardQuery' `
        -Sheet 'Services' -QueryTable 'OperationalDashboardQueryTable' `
        -CommandText 'let Source = Excel.CurrentWorkbook(){[Name="ServiceHealth"]}[Content] in Source' `
        -Description 'Refresh metadata for Excel-compatible applications; PSWriteOffice does not execute Power Query.' `
        -RefreshOnOpen
}

$checks = Test-OfficeDocument -Path $path
$accessibility = Test-OfficeDocument -Path $path -Check Accessibility
$comments = @(Get-OfficeExcelComment -Path $path -Kind All)
$threadedComments = @(Get-OfficeExcelComment -Path $path -Kind Threaded)
$dataModel = Get-OfficeExcelDataModel -Path $path
$repair = Repair-OfficeDocument -Path $path -PassThru
$info = Get-OfficeDocumentInfo -Path $path -Detailed
$native = $info.Details

Write-Host "Workbook saved to $path"
Write-Host "Workbook summary: $($info.SheetCount) sheets, $($native.ChartPartCount) chart parts, $($native.ConnectionPartCount) connection parts"
Write-Host "Workbook checks: valid=$($checks.IsValid), accessible=$($accessibility.IsValid), comments=$($comments.Count), threaded=$($threadedComments.Count), queries=$($dataModel.HasDataModelOrQueries), repairActions=$($repair.ActionCount)"
$checks.NativeReports

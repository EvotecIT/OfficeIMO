Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$path = Join-Path $documents 'Excel-ModifyExistingTables.xlsx'

$initialRows = @(
    [PSCustomObject]@{ Service = 'Identity'; Status = 'Ready'; Owner = 'IAM' }
    [PSCustomObject]@{ Service = 'Messaging'; Status = 'Watching'; Owner = 'Collaboration' }
)

New-OfficeExcel -Path $path {
    ExcelSheet 'Readiness' {
        ExcelTable -Data $initialRows -TableName 'ServiceReadiness' -AutoFit
    }
    ExcelSheet 'Notes' {
        Set-OfficeExcelCell -Address A1 -Value 'This workbook is modified after it is created.'
    }
}

# Second pass: append rows to the named table without rebuilding the workbook through the DSL.
$workbook = Get-OfficeExcel -Path $path
try {
    $workbook |
        Add-OfficeExcelTableRow -Sheet Readiness -TableName ServiceReadiness -InputObject ([PSCustomObject]@{
            Service = 'File Services'
            Status  = 'Ready'
            Owner   = 'Storage'
        }) -PassThru

    $moreRows = @(
        [PSCustomObject]@{ Service = 'Network'; Status = 'Investigating'; Owner = 'Platform' }
        [PSCustomObject]@{ Service = 'Monitoring'; Status = 'Ready'; Owner = 'SRE' }
    )

    $workbook |
        Add-OfficeExcelTableRow -Sheet Readiness -TableName ServiceReadiness -InputObject $moreRows
} finally {
    Close-OfficeExcel -Document $workbook -Save
}

Write-Host "Updated Excel workbook saved to $path"
Write-Host 'Tables:'
Get-OfficeExcelTable -Path $path -Sheet Readiness -Name ServiceReadiness |
    Format-Table

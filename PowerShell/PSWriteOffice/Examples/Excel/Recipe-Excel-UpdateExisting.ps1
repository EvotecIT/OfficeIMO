$path = '.\Excel-Updated-Existing.xlsx'
$rows = @(
    [pscustomobject]@{ Service = 'Identity'; Status = 'Draft'; Owner = 'IAM' }
    [pscustomobject]@{ Service = 'Messaging'; Status = 'Draft'; Owner = 'Collaboration' }
)

ExcelNew -Path $path {
    ExcelSheet 'Readiness' {
        ExcelTable -Data $rows -TableName 'Readiness' -AutoFit
    }
}

Update-OfficeExcelText -Path $path -Sheet 'Readiness' -OldValue 'Draft' -NewValue 'Ready'

Edit-OfficeExcelRow -Path $path -Sheet 'Readiness' -ScriptBlock {
    param($row)

    if ($row.CellByHeader('Service').Value -eq 'Messaging') {
        $row.Set('Owner', 'Productivity')
    }
}

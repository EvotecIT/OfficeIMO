$path = '.\Excel-Updated-Existing.xlsx'
$rows = @(
    [pscustomobject]@{ Service = 'Identity'; Status = 'Draft'; Owner = 'IAM' }
    [pscustomobject]@{ Service = 'Messaging'; Status = 'Draft'; Owner = 'Collaboration' }
)

New-OfficeExcel -Path $path {
    ExcelSheet 'Readiness' {
        ExcelTable -InputObject $rows -Name 'Readiness'
    }
}

Update-OfficeExcelText -Path $path -Sheet 'Readiness' -Pattern 'Draft' -Replacement 'Ready'

Edit-OfficeExcelRow -Path $path -Sheet 'Readiness' -ScriptBlock {
    param($row)

    if ($row.CellByHeader('Service').Value -eq 'Messaging') {
        $row.Set('Owner', 'Productivity')
    }
}

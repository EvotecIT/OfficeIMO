$path = '.\Excel-Read-And-Filter.xlsx'
$rows = @(
    [pscustomobject]@{ Service = 'Identity'; Owner = 'IAM'; Incidents = 1; Status = 'Ready' }
    [pscustomobject]@{ Service = 'Messaging'; Owner = 'Collaboration'; Incidents = 5; Status = 'Review' }
    [pscustomobject]@{ Service = 'Files'; Owner = 'Storage'; Incidents = 0; Status = 'Ready' }
)

New-OfficeExcel -Path $path {
    ExcelSheet 'Services' {
        ExcelTable -Data $rows -Name 'ServiceHealth' -AutoFit
    }
}

Import-OfficeExcel -Path $path -WorksheetName 'Services' |
    Where-Object { $_.Status -eq 'Review' -or $_.Incidents -ge 3 } |
    Select-Object Service, Owner, Incidents, Status

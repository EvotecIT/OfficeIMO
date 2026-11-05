$path = '.\Daily-Orders.xlsx'
$morning = @(
    [pscustomobject]@{ Order = 'SO-1001'; Owner = 'Sales'; Status = 'Ready' }
    [pscustomobject]@{ Order = 'SO-1002'; Owner = 'Sales'; Status = 'Review' }
)
$afternoon = @(
    [pscustomobject]@{ Order = 'SO-1003'; Owner = 'Support'; Status = 'Ready' }
)

$morning | Export-OfficeExcel -Path $path -Sheet 'Orders' -TableName 'Orders'
$afternoon | Export-OfficeExcel -Path $path -Sheet 'Orders' -TableName 'Orders' -Append

$summary = @(
    [pscustomobject]@{ Status = 'Ready'; Count = 2 }
    [pscustomobject]@{ Status = 'Review'; Count = 1 }
)
$summary | Export-OfficeExcel -Path $path -Sheet 'Summary' -TableName 'Summary' -ReplaceSheet

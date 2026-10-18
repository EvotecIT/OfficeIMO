$orders = @(
    [pscustomobject]@{ Order = 'SO-1001'; Customer = 'Northwind'; Total = 1250.50; Due = [datetime]'2026-09-05' }
    [pscustomobject]@{ Order = 'SO-1002'; Customer = 'Contoso'; Total = 840.00; Due = [datetime]'2026-09-08' }
)

$orders | Export-OfficeExcel `
    -Path '.\Orders.xlsx' `
    -WorksheetName 'Orders' `
    -TableName 'Orders' `
    -CurrencyColumn Total `
    -DateColumn Due `
    -AutoFit `
    -FreezeTopRow

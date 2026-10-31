$orders = @(
    [pscustomobject]@{ Order = 'SO-1001'; Customer = 'Northwind'; Total = 1250.50; Due = [datetime]'2026-09-05' }
    [pscustomobject]@{ Order = 'SO-1002'; Customer = 'Contoso'; Total = 840.00; Due = [datetime]'2026-09-08' }
)

# Styled table, frozen header, fitted columns and a date format for Due come by default.
$orders | Export-OfficeExcel -Path .\Orders.xlsx -Sheet Orders -TableName Orders -ColumnFormat @{ Total = 'Currency' }

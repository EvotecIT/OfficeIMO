---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportKpi
## SYNOPSIS
Adds a row of KPI cards.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportKpi [-InputObject] <Object> [-PerRow <int>] [-LabelFillColor <OfficeColor>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a row of KPI cards.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportKpi ([ordered]@{ Uptime = '99.97%'; Incidents = 4; Deployments = 37 })
    }
}
```

Shows three headline figures as KPI cards.

### EXAMPLE 2
```powershell
$orders = Import-OfficeCsv -Path .\Data\Q3-Orders.csv
$kpis = [ordered]@{
    Revenue           = ($orders | Measure-Object Amount -Sum).Sum
    Orders            = $orders.Count
    'Average order'   = [math]::Round(($orders | Measure-Object Amount -Average).Average, 2)
    'Active customers' = ($orders | Select-Object -ExpandProperty Customer -Unique).Count
}
New-OfficeExcel -Path .\Reports\Q3-Sales-Summary.xlsx {
    ExcelReportSheet 'Summary' {
        ExcelReportTitle 'Q3 sales' 'July to September 2026'
        ExcelReportKpi $kpis -PerRow 4
    }
}
```

Calculates the sales KPIs from the order export and shows them four to a row.

## PARAMETERS

### -InputObject
Objects, dictionaries or native tabular data supplied directly or through the pipeline.

```yaml
Type: Object
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -LabelFillColor
Optional fill color for KPI labels.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PerRow
Number of KPI cards per rendered row.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.Object`

## OUTPUTS

- `None`

## RELATED LINKS

- None

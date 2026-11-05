---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelChart
## SYNOPSIS
Applies a built-in style and color preset to a chart.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelChart [-Chart] <ExcelChart> [-Style <int>] [-ColorStyle <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Applies a built-in style and color preset to a chart.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Regions.xlsx {
    ExcelSheet 'Summary' {
        ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
        ExcelChart -Table RegionSummary -Address F2 -Type Column -PassThru | Set-OfficeExcelChart -Style 3
    }
}
```

Applies built-in chart style 3 to a chart returned by ExcelChart.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Operations.xlsx {
    ExcelSheet 'Dashboard' {
        $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Site-Uptime.csv) -Category Site -Value Uptime -Address A1 -Type Column -PassThru
        Set-OfficeExcelChart -Chart $chart -Style 6 -ColorStyle 14
    }
}
```

Uses a built-in style and colour preset for a chart that sits on a branded dashboard.

## PARAMETERS

### -Chart
Chart to update.

```yaml
Type: ExcelChart
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ColorStyle
Built-in chart color-style identifier.

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

### -Style
Built-in chart style identifier.

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

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `None`

## RELATED LINKS

- None

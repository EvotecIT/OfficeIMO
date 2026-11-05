---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelChartLegend
## SYNOPSIS
Shows, hides, positions and styles the chart legend.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelChartLegend [-Chart] <ExcelChart> [-Position <OfficeChartLegendPosition>] [-Hidden] [-Overlay] [-Font <string>] [-Size <OfficeLength>] [-Color <OfficeColor>] [-Bold] [-Italic] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Shows, hides, positions and styles the chart legend.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Helpdesk-Trend.xlsx {
    ExcelSheet 'Trend' {
        ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Ticket-Trend.csv) -Category Week -Value Opened, Closed -Type Line -Address A1 -PassThru | ExcelChartLegend -Position Bottom
    }
}
```

Places the legend of a line chart under the plot area.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Availability.xlsx {
    ExcelSheet 'Uptime' {
        $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Site-Uptime.csv) -Category Site -Value Uptime -Address A1 -Type Column -Title 'Uptime by site' -PassThru
        ExcelChartLegend -Chart $chart -Hidden
    }
}
```

Removes the legend where the chart title already names the only series.

## PARAMETERS

### -Bold
Applies bold text. An explicit false value clears bold; omission preserves the native value.

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

### -Color
Text or rule colour expressed as a literal OfficeColor or theme role.

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

### -Font
Font family name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
Hide or show the column.

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

### -Italic
Applies italic text. An explicit false value clears italic; omission preserves the native value.

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

### -Overlay
Overlay the legend on the chart area.

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

### -Position
Legend position.

```yaml
Type: OfficeChartLegendPosition
Parameter Sets: Default
Aliases: None
Possible values: Bottom, TopRight, Left, Right, Top

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Font size, interpreted as points for numbers; explicit physical units are accepted.

```yaml
Type: OfficeLength
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

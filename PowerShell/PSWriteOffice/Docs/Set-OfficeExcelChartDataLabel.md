---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelChartDataLabel
## SYNOPSIS
Shows and styles chart data labels.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelChartDataLabel [-Chart] <ExcelChart> [-ShowValue] [-ShowCategoryName] [-ShowSeriesName] [-ShowPercent] [-ShowLegendKey] [-Position <OfficeChartDataLabelPosition>] [-NumberFormat <string>] [-SourceLinked] [-Font <string>] [-Size <OfficeLength>] [-Color <OfficeColor>] [-Bold] [-Italic] [-FillColor <OfficeColor>] [-NoFill] [-LineColor <OfficeColor>] [-LineWidth <OfficeLength>] [-NoLine] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Shows and styles chart data labels.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Regions.xlsx {
    ExcelSheet 'Summary' {
        ExcelTable (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Name RegionSummary
        ExcelChart -Table RegionSummary -Address F2 -Type Column -PassThru | ExcelChartDataLabel -ShowValue
    }
}
```

Shows each bar's value on the chart.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Cloud-Spend.xlsx {
    ExcelSheet 'Spend' {
        $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Spend-By-Service.csv) -Category Service -Value Cost -Address A1 -Type Pie -Title 'Cloud spend by service' -PassThru
        ExcelChartDataLabel -Chart $chart -ShowCategoryName -ShowPercent -Position OutsideEnd -NumberFormat '0%'
    }
}
```

Labels each slice of a pie chart with its category and percentage outside the slice.

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

### -FillColor
Optional label fill color. Named colors and hexadecimal values are accepted.

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

### -LineColor
Optional label line color. Named colors and hexadecimal values are accepted.

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

### -LineWidth
Physical line width; numbers are interpreted as points.

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

### -NoFill
Remove label fill.

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

### -NoLine
Remove label border.

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

### -NumberFormat
Optional number format code.

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
Optional data label position.

```yaml
Type: OfficeChartDataLabelPosition
Parameter Sets: Default
Aliases: None
Possible values: BestFit, Bottom, Center, InsideBase, InsideEnd, Left, OutsideEnd, Right, Top

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowCategoryName
Show category names in labels.

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

### -ShowLegendKey
Show legend keys in labels.

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

### -ShowPercent
Show percentages in labels.

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

### -ShowSeriesName
Show series names in labels.

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

### -ShowValue
Show values in labels.

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

### -SourceLinked
Keep number formatting linked to the source cells.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `None`

## RELATED LINKS

- None

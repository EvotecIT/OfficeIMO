---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelChartTrendline
## SYNOPSIS
Adds or replaces a series trendline.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartTrendline [-Chart] <ExcelChart> -Index <int> [-Type <OfficeChartTrendlineType>] [-Order <Int32>] [-Period <Int32>] [-Forward <Double>] [-Backward <Double>] [-Intercept <Double>] [-DisplayEquation] [-DisplayRSquared] [-LineColor <OfficeColor>] [-LineWidth <OfficeLength>] [-PassThru] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartTrendline [-Chart] <ExcelChart> -Name <string> [-Type <OfficeChartTrendlineType>] [-Order <Int32>] [-Period <Int32>] [-Forward <Double>] [-Backward <Double>] [-Intercept <Double>] [-DisplayEquation] [-DisplayRSquared] [-LineColor <OfficeColor>] [-LineWidth <OfficeLength>] [-PassThru] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Adds or replaces a series trendline.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Growth.xlsx {
    ExcelSheet 'Users' {
        ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Monthly-Users.csv) -Category Month -Value ActiveUsers -Type Line -Address A1 -PassThru | ExcelChartTrendline -Index 1 -Type Linear
    }
}
```

Adds a straight trendline to the first series.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Storage-Forecast.xlsx {
    ExcelSheet 'Storage' {
        $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Storage-Used.csv) -Category Month -Value UsedTB -Type Line -Address A1 -Title 'Storage used (TB)' -PassThru
        ExcelChartTrendline -Chart $chart -Name UsedTB -Type Linear -Forward 3 -DisplayRSquared
    }
}
```

Fits a trendline to monthly storage use, projects it three months forward and shows the R-squared value on the chart.

## PARAMETERS

### -Backward
Backward forecast units.

```yaml
Type: Double
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CaseSensitive
Matches names and text using ordinal case sensitivity. Omission ignores case.

```yaml
Type: SwitchParameter
Parameter Sets: Name
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
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DisplayEquation
Display the trendline equation.

```yaml
Type: SwitchParameter
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisplayRSquared
Display the R-squared value.

```yaml
Type: SwitchParameter
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Forward
Forward forecast units.

```yaml
Type: Double
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Index
One-based index identifying the target item.

```yaml
Type: Int32
Parameter Sets: Index
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Intercept
Trendline intercept.

```yaml
Type: Double
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineColor
Trendline line color. Named colors and hexadecimal values are accepted.

```yaml
Type: OfficeColor
Parameter Sets: Index, Name
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
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Named range name.

```yaml
Type: String
Parameter Sets: Name
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Order
Polynomial order.

```yaml
Type: Int32
Parameter Sets: Index, Name
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
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Period
Moving-average period.

```yaml
Type: Int32
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Type
Trendline type.

```yaml
Type: OfficeChartTrendlineType
Parameter Sets: Index, Name
Aliases: None
Possible values: Exponential, Linear, Logarithmic, MovingAverage, Polynomial, Power

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

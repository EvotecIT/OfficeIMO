---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartTrendline
## SYNOPSIS
Adds or replaces an Excel chart series trendline.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartTrendline -Chart <ExcelChart> -SeriesIndex <int> -Type <OfficeChartTrendlineType> [-Order <Int32>] [-Period <Int32>] [-Forward <Double>] [-Backward <Double>] [-Intercept <Double>] [-DisplayEquation] [-DisplayRSquared] [-LineColor <string>] [-LineWidthPoints <Double>] [-PassThru] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartTrendline -Chart <ExcelChart> -SeriesName <string> -Type <OfficeChartTrendlineType> [-IgnoreCase <bool>] [-Order <Int32>] [-Period <Int32>] [-Forward <Double>] [-Backward <Double>] [-Intercept <Double>] [-DisplayEquation] [-DisplayRSquared] [-LineColor <string>] [-LineWidthPoints <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds or replaces an Excel chart series trendline.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $chart | Set-OfficeExcelChartTrendline -SeriesIndex 0 -Type Polynomial -Order 2 -DisplayEquation -DisplayRSquared
```

Adds a polynomial trendline to the first series.

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

### -Chart
Chart to update.

```yaml
Type: ExcelChart
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: True
Position: named
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

### -IgnoreCase
Ignore case when matching series name.

```yaml
Type: Boolean
Parameter Sets: Name
Aliases: None
Possible values:

Required: False
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
Type: String
Parameter Sets: Index, Name
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineWidthPoints
Trendline line width in points.

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
Emit the object created or changed by the command.

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

### -SeriesIndex
Zero-based series index.

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

### -SeriesName
Series name.

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

### -Type
Trendline type.

```yaml
Type: OfficeChartTrendlineType
Parameter Sets: Index, Name
Aliases: None
Possible values: Exponential, Linear, Logarithmic, MovingAverage, Polynomial, Power

Required: True
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

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None

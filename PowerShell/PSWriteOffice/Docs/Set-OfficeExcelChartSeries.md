---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelChartSeries
## SYNOPSIS
Configures Excel chart series colors, line style, and markers.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartSeries -Chart <ExcelChart> -SeriesIndex <int> [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <Double>] [-MarkerStyle <OfficeChartMarkerShape>] [-MarkerSize <Int32>] [-MarkerFillColor <string>] [-MarkerLineColor <string>] [-MarkerLineWidthPoints <Double>] [-PassThru] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartSeries -Chart <ExcelChart> -SeriesName <string> [-IgnoreCase <bool>] [-FillColor <string>] [-LineColor <string>] [-LineWidthPoints <Double>] [-MarkerStyle <OfficeChartMarkerShape>] [-MarkerSize <Int32>] [-MarkerFillColor <string>] [-MarkerLineColor <string>] [-MarkerLineWidthPoints <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Configures Excel chart series colors, line style, and markers.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $chart | Set-OfficeExcelChartSeries -SeriesName 'Revenue' -FillColor '#4472C4' -LineColor '#1F4E79' -MarkerStyle Circle -MarkerSize 6
```

Applies fill, line, and marker settings to the Revenue series.

## PARAMETERS

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

### -FillColor
Series fill color. Named colors and hexadecimal values are accepted.

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

### -LineColor
Series line color. Named colors and hexadecimal values are accepted.

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
Series line width in points.

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

### -MarkerFillColor
Marker fill color. Named colors and hexadecimal values are accepted.

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

### -MarkerLineColor
Marker line color. Named colors and hexadecimal values are accepted.

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

### -MarkerLineWidthPoints
Marker line width in points.

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

### -MarkerSize
Marker size.

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

### -MarkerStyle
Marker style name.

```yaml
Type: OfficeChartMarkerShape
Parameter Sets: Index, Name
Aliases: None
Possible values: Auto, Circle, Dash, Diamond, Dot, None, Picture, Plus, Square, Star, Triangle, X

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelChart`

## OUTPUTS

- `OfficeIMO.Excel.ExcelChart`

## RELATED LINKS

- None

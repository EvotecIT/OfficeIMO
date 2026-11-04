---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelChartSeries
## SYNOPSIS
Styles a chart series or one of its data points.

## SYNTAX
### Index (Default)
```powershell
Set-OfficeExcelChartSeries [-Chart] <ExcelChart> -Index <int> [-Point <UInt32>] [-FillColor <OfficeColor>] [-LineColor <OfficeColor>] [-LineWidth <OfficeLength>] [-MarkerStyle <OfficeChartMarkerShape>] [-MarkerSize <Int32>] [-MarkerFillColor <OfficeColor>] [-MarkerLineColor <OfficeColor>] [-MarkerLineWidth <OfficeLength>] [-PassThru] [<CommonParameters>]
```

### Name
```powershell
Set-OfficeExcelChartSeries [-Chart] <ExcelChart> -Name <string> [-Point <UInt32>] [-FillColor <OfficeColor>] [-LineColor <OfficeColor>] [-LineWidth <OfficeLength>] [-MarkerStyle <OfficeChartMarkerShape>] [-MarkerSize <Int32>] [-MarkerFillColor <OfficeColor>] [-MarkerLineColor <OfficeColor>] [-MarkerLineWidth <OfficeLength>] [-PassThru] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Styles a chart series or one of its data points.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeExcel -Path .\Reports\Helpdesk-Trend.xlsx {
    ExcelSheet 'Trend' {
        ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Ticket-Trend.csv) -Category Week -Value Opened, Closed -Type Line -Address A1 -PassThru | ExcelChartSeries -Index 2 -LineColor Success
    }
}
```

Draws the second series, Closed, in the theme's success colour.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Regions.xlsx {
    ExcelSheet 'Summary' {
        $chart = ExcelChart -InputObject (Import-OfficeCsv -Path .\Data\Region-Revenue.csv) -Category Region -Value Revenue -Address A1 -Type Column -PassThru
        ExcelChartSeries -Chart $chart -Name Revenue -Point 3 -FillColor Danger
    }
}
```

Colours the bar of the region that missed its target in the danger colour and leaves the others on the theme palette.

## PARAMETERS

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

### -FillColor
Series fill color. Named colors and hexadecimal values are accepted.

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

### -LineColor
Series line color. Named colors and hexadecimal values are accepted.

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

### -MarkerFillColor
Marker fill color. Named colors and hexadecimal values are accepted.

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

### -MarkerLineColor
Marker line color. Named colors and hexadecimal values are accepted.

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

### -MarkerLineWidth
Physical width of the marker outline; omission preserves the existing width.

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

### -Point
One-based point index within the selected chart series.

```yaml
Type: UInt32
Parameter Sets: Index, Name
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

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelSparkline
## SYNOPSIS
Adds sparklines to a worksheet.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelSparkline -DataRange <string> -LocationRange <string> [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Type <ExcelSparklineType>] [-ShowMarkers] [-ShowHighLow] [-ShowFirstLast] [-ShowNegative] [-ShowAxis] [-SeriesColor <OfficeColor>] [-AxisColor <OfficeColor>] [-NegativeColor <OfficeColor>] [-MarkerColor <OfficeColor>] [-HighColor <OfficeColor>] [-LowColor <OfficeColor>] [-FirstColor <OfficeColor>] [-LastColor <OfficeColor>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds sparklines to a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
    ExcelSheet 'Departments' {
        ExcelSparkline -DataRange B2:M20 -LocationRange N2:N20
    }
}
```

Draws a sparkline for each row of monthly figures in the column next to them.

### EXAMPLE 2
```powershell
New-OfficeExcel -Path .\Reports\Site-Incidents.xlsx {
    ExcelSheet 'Incidents' {
        ExcelTable (Import-OfficeCsv -Path .\Data\Incidents-By-Week.csv) -Name Incidents
        ExcelSparkline -DataRange B2:M9 -LocationRange N2:N9 -Type Column -ShowHighLow
    }
}
```

Adds column sparklines with the highest and lowest weeks marked, so each site's pattern is visible at a glance.

## PARAMETERS

### -AxisColor
Axis color. Named colors and hexadecimal values are accepted.

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

### -DataRange
A1 range containing source data (e.g., "B2:M2").

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to operate on outside the DSL context.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FirstColor
First point color. Named colors and hexadecimal values are accepted.

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

### -HighColor
High point color. Named colors and hexadecimal values are accepted.

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

### -LastColor
Last point color. Named colors and hexadecimal values are accepted.

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

### -LocationRange
A1 range where the sparklines will be placed (e.g., "N2:N2").

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LowColor
Low point color. Named colors and hexadecimal values are accepted.

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

### -MarkerColor
Markers color. Named colors and hexadecimal values are accepted.

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

### -NegativeColor
Negative point color. Named colors and hexadecimal values are accepted.

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
Emit the worksheet after adding sparklines.

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

### -SeriesColor
Series color. Named colors and hexadecimal values are accepted.

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

### -Sheet
Worksheet name when using Document.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowAxis
Show axis.

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

### -ShowFirstLast
Show first/last points.

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

### -ShowHighLow
Show high/low points.

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

### -ShowMarkers
Show markers for each point.

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

### -ShowNegative
Show negative points.

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

### -Type
Sparkline type (Line, Column, WinLoss).

```yaml
Type: ExcelSparklineType
Parameter Sets: Default
Aliases: None
Possible values: Line, Column, Stacked

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None

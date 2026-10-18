---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelSparkline
## SYNOPSIS
Adds sparklines to a worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelSparkline -DataRange <string> -LocationRange <string> [-Type <ExcelSparklineType>] [-ShowMarkers] [-ShowHighLow] [-ShowFirstLast] [-ShowNegative] [-ShowAxis] [-SeriesColor <string>] [-AxisColor <string>] [-NegativeColor <string>] [-MarkersColor <string>] [-HighColor <string>] [-LowColor <string>] [-FirstColor <string>] [-LastColor <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelSparkline -Document <ExcelDocument> -DataRange <string> -LocationRange <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Type <ExcelSparklineType>] [-ShowMarkers] [-ShowHighLow] [-ShowFirstLast] [-ShowNegative] [-ShowAxis] [-SeriesColor <string>] [-AxisColor <string>] [-NegativeColor <string>] [-MarkersColor <string>] [-HighColor <string>] [-LowColor <string>] [-FirstColor <string>] [-LastColor <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds sparklines to a worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelSparkline -DataRange 'B2:M2' -LocationRange 'N2' }
```

Creates a line sparkline in N2 using B2:M2.

## PARAMETERS

### -AxisColor
Axis color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FirstColor
First point color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Context, Document
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
Type: String
Parameter Sets: Context, Document
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
Type: String
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarkersColor
Markers color. Named colors and hexadecimal values are accepted.

```yaml
Type: String
Parameter Sets: Context, Document
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
Type: String
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Type: String
Parameter Sets: Context, Document
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
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) when using Document.

```yaml
Type: Int32
Parameter Sets: Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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
Parameter Sets: Context, Document
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

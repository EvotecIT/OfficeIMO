---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelConditionalIconSet
## SYNOPSIS
Adds an icon set conditional format to a range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelConditionalIconSet [[-Range] <string>] [-HeaderName <string>] [-TableName <string>] [-PivotTableName <string>] [-PivotWholeTable] [-HeaderRow <int>] [-IncludeHeader] [-IconSet <ExcelIconSet>] [-ShowValue <bool>] [-Reverse <bool>] [-PercentThresholds <double[]>] [-NumberThresholds <double[]>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelConditionalIconSet [[-Range] <string>] -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-HeaderName <string>] [-TableName <string>] [-PivotTableName <string>] [-PivotWholeTable] [-HeaderRow <int>] [-IncludeHeader] [-IconSet <ExcelIconSet>] [-ShowValue <bool>] [-Reverse <bool>] [-PercentThresholds <double[]>] [-NumberThresholds <double[]>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds an icon set conditional format to a range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelConditionalIconSet -Range 'E2:E50' -IconSet ThreeTrafficLights1 }
```

Applies a traffic-light icon set.

## PARAMETERS

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

### -HeaderName
Header or table column name used to resolve the target range.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: ColumnName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
Worksheet header row used when resolving HeaderName without a table. Use 0 for the first row of the used range.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IconSet
Icon set to apply.

```yaml
Type: ExcelIconSet
Parameter Sets: Context, Document
Aliases: None
Possible values: ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeader
Include the header cell in the resolved range.

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

### -NumberThresholds
Number thresholds matching the icon count.

```yaml
Type: Double[]
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
Emit the range after applying the format.

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

### -PercentThresholds
Percent thresholds (0..100) matching the icon count.

```yaml
Type: Double[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotTableName
Pivot table name used to resolve the target range.

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

### -PivotWholeTable
Use the full pivot output range instead of the default data body range.

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

### -Range
A1 range to format.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reverse
Reverse the icon order.

```yaml
Type: Boolean
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

### -ShowValue
Show the underlying values.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableName
Optional table name for header-based range resolution.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None

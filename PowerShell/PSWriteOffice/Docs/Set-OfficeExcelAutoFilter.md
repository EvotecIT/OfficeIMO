---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelAutoFilter
## SYNOPSIS
Applies a friendly AutoFilter condition by header name.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelAutoFilter [-Header] <string> [-Range <string>] [-Value <string[]>] [-Contains <string>] [-DoesNotContain <string>] [-StartsWith <string>] [-EndsWith <string>] [-GreaterThanOrEqual <Double>] [-LessThanOrEqual <Double>] [-NotEqual <Double>] [-Between <double[]>] [-NotBetween <double[]>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelAutoFilter [-Header] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-Value <string[]>] [-Contains <string>] [-DoesNotContain <string>] [-StartsWith <string>] [-EndsWith <string>] [-GreaterThanOrEqual <Double>] [-LessThanOrEqual <Double>] [-NotEqual <Double>] [-Between <double[]>] [-NotBetween <double[]>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Applies a friendly AutoFilter condition by header name.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelAutoFilter -Range A1:D200 -Header Status -Value Open,Hold }
```

Ensures an AutoFilter range exists and filters the Status column.

## PARAMETERS

### -Between
Inclusive numeric range condition. Provide exactly two values: minimum, maximum.

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

### -Contains
Text that the column value must contain.

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

### -DoesNotContain
Text that the column value must not contain.

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

### -EndsWith
Text that the column value must end with.

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

### -GreaterThanOrEqual
Numeric greater-than-or-equal condition.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header name to filter.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: HeaderName, ColumnName
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LessThanOrEqual
Numeric less-than-or-equal condition.

```yaml
Type: Double
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NotBetween
Outside numeric range condition. Provide exactly two values: minimum, maximum.

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

### -NotEqual
Numeric not-equal condition.

```yaml
Type: Double
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
Emit the worksheet after applying the filter.

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
Optional A1 AutoFilter range to create or replace before applying the condition.

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

### -StartsWith
Text that the column value must start with.

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

### -Value
Allowed values for an equals filter.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: Values
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

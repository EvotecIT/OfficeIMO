---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelValidationList
## SYNOPSIS
Adds a list-based data validation to a worksheet range.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelValidationList [[-Range] <string>] [-Values] <string[]> [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [-AllowBlank <bool>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelValidationList [[-Range] <string>] [-Values] <string[]> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [-AllowBlank <bool>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
When invoked inside the Excel DSL, applies to the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelValidationList -Range 'C2:C50' -Values 'New','In Progress','Done' }
```

Restricts column C to the provided values.

## PARAMETERS

### -AllowBlank
Allow blank values.

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

### -PassThru
Emit the range string after applying validation.

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
Target range in A1 notation.

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

### -Values
Allowed values for the dropdown list.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 1
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

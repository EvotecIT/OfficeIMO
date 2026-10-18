---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelConditionalFormatting
## SYNOPSIS
Gets conditional formatting rules from one or more Excel worksheets.

## SYNTAX
### Context (Default)
```powershell
Get-OfficeExcelConditionalFormatting [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelConditionalFormatting [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelConditionalFormatting -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-HeaderName <string>] [-TableName <string>] [-HeaderRow <int>] [-IncludeHeader] [<CommonParameters>]
```

## DESCRIPTION
Gets conditional formatting rules from one or more Excel worksheets.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeExcelConditionalFormatting -Path .\Report.xlsx -Sheet Data |
                Select-Object -Property SheetName, Range, Type, Operator, Formulas
```

Returns rule metadata that can be filtered, exported, or used before clearing stale rules.

## PARAMETERS

### -Document
Workbook to inspect outside the DSL context.

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
Header or table column name used to resolve the range filter.

```yaml
Type: String
Parameter Sets: Context, Path, Document
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
Parameter Sets: Context, Path, Document
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
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path to inspect.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range
Optional A1 range filter.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name to inspect. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) to inspect. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
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
Parameter Sets: Context, Path, Document
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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None

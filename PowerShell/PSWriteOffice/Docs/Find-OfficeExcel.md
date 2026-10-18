---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Find-OfficeExcel
## SYNOPSIS
Finds text in worksheet values.

## SYNTAX
### Path (Default)
```powershell
Find-OfficeExcel [-Path] <string> [-Text] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-CaseSensitive] [-Regex] [-Exact] [<CommonParameters>]
```

### Document
```powershell
Find-OfficeExcel [-Text] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-CaseSensitive] [-Regex] [-Exact] [<CommonParameters>]
```

## DESCRIPTION
Finds text in worksheet values.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $matches = Find-OfficeExcel -Path .\Report.xlsx -Text Ready -Sheet Summary
$matches |
    Select-Object -Property Sheet, Address, Value |
    Export-Csv -Path .\ReadyCells.csv -NoTypeInformation
```

Returns matching cells with sheet, address, row, column, and value metadata for review or proof.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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

### -Exact
Require an exact cell text match instead of substring matching.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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
A1 range to search. Defaults to each selected worksheet's used range.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Regex
Treat -Text as a regular expression.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name. Defaults to all sheets for path/document use and current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index when using a workbook object or path.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text or pattern to find.

```yaml
Type: String
Parameter Sets: Path, Document
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

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelConditionalFormat
## SYNOPSIS
Gets native conditional formatting snapshots, with worksheet and source-path metadata.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelConditionalFormat [-Range <string>] [-Header <string>] [-Table <string>] [-HeaderRow <int>] [-IncludeHeader] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelConditionalFormat [-Path] <string[]> [-Range <string>] [-Header <string>] [-Table <string>] [-HeaderRow <int>] [-IncludeHeader] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelConditionalFormat -LiteralPath <string[]> [-Range <string>] [-Header <string>] [-Table <string>] [-HeaderRow <int>] [-IncludeHeader] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

## DESCRIPTION
Gets native conditional formatting snapshots, with worksheet and source-path metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelConditionalFormat -Path .\Report.xlsx -Sheet Data -Table Sales -Header Status
```


## PARAMETERS

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Header
Header or table column name used to resolve the range filter.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderRow
Worksheet header row; zero uses the first row of the used range.

```yaml
Type: Int32
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeHeader
Include the header cell when resolving a named column.

```yaml
Type: SwitchParameter
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal workbook paths. FileInfo.FullName and provider PSPath bind here.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: FullName, PSPath
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Path
Workbook paths, including filesystem wildcard patterns.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Range
A1 range filter; Header takes precedence when both are supplied.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet names, wildcard patterns, one-based positions or native worksheet objects.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Document, Path, LiteralPath
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Table
Table name for header-based range resolution.

```yaml
Type: String
Parameter Sets: Document, Path, LiteralPath
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

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `OfficeIMO.Excel.ExcelConditionalFormattingInfo`

## RELATED LINKS

- None

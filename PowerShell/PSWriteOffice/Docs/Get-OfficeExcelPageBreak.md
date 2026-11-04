---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelPageBreak
## SYNOPSIS
Lists manual print page breaks without changing a workbook.

## SYNTAX
### Document (Default)
```powershell
Get-OfficeExcelPageBreak [-Row] [-Column] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### Path
```powershell
Get-OfficeExcelPageBreak [-Path] <string[]> [-Row] [-Column] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

### LiteralPath
```powershell
Get-OfficeExcelPageBreak -LiteralPath <string[]> [-Row] [-Column] [-Sheet <OfficeSheetSelector[]>] [<CommonParameters>]
```

## DESCRIPTION
Lists manual print page breaks without changing a workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-ChildItem ./Reports -Filter *.xlsx | Get-OfficeExcelPageBreak -Sheet 'Report*'
```


## PARAMETERS

### -Column
Include column breaks; neither switch or both switches includes both kinds.

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

### -Row
Include row breaks; neither switch or both switches includes both kinds.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `System.String[]`
- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `PSWriteOffice.Models.Excel.ExcelPageBreakInfo`: A manual print page-break snapshot that remains readable after its source workbook is closed.

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelSummary
## SYNOPSIS
Gets a compact structural summary of an Excel workbook.

## SYNTAX
### Path (Default)
```powershell
Get-OfficeExcelSummary [-Path] <string> [-IncludeSheets] [-IncludeSchema] [<CommonParameters>]
```

### Document
```powershell
Get-OfficeExcelSummary -Document <ExcelDocument> [-IncludeSheets] [-IncludeSchema] [<CommonParameters>]
```

## DESCRIPTION
Gets a compact structural summary of an Excel workbook.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $summary = Get-OfficeExcelSummary -Path .\report.xlsx -IncludeSheets
$summary |
    Select-Object -Property SheetCount, TableCount, ChartCount, PivotTableCount
$summary.Sheets |
    Select-Object -Property Name, State, UsedRange
```

Returns workbook-level counts plus per-sheet tables, charts, pivots, links, comments, and used ranges.

## PARAMETERS

### -Document
Workbook to inspect.

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

### -IncludeSchema
Include OfficeIMO inspection snapshot details for schema discovery.

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

### -IncludeSheets
Include per-sheet details in the returned object.

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
Path to the workbook.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Compare-OfficeExcelWorkbook
## SYNOPSIS
Compares two workbooks by sheets, cells, formulas, styles, tables, comments, names, and worksheet metadata.

## SYNTAX
### Path (Default)
```powershell
Compare-OfficeExcelWorkbook [-Path] <string> [-DifferencePath] <string> [-MaxDifferences <int>] [-SkipCells] [-SkipCellStyles] [-SkipNamedRanges] [-SkipTables] [-SkipWorksheetMetadata] [-SkipComments] [<CommonParameters>]
```

### Document
```powershell
Compare-OfficeExcelWorkbook -Document <ExcelDocument> -DifferenceDocument <ExcelDocument> [-MaxDifferences <int>] [-SkipCells] [-SkipCellStyles] [-SkipNamedRanges] [-SkipTables] [-SkipWorksheetMetadata] [-SkipComments] [<CommonParameters>]
```

## DESCRIPTION
Compares two workbooks by sheets, cells, formulas, styles, tables, comments, names, and worksheet metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $diff = Compare-OfficeExcelWorkbook -Path .\Expected.xlsx -DifferencePath .\Actual.xlsx -MaxDifferences 500
if (-not $diff.AreEqual) {
    $diff.Differences |
        Sort-Object Category,SheetName,Address |
        Format-Table Category,SheetName,Address,Message,LeftValue,RightValue
}
```

Uses OfficeIMO's reusable workbook diff engine and includes structural metadata by default, not just visible cell values.

## PARAMETERS

### -DifferenceDocument
Workbook document to compare against.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferencePath
Workbook path to compare against.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook document.

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

### -MaxDifferences
Maximum number of differences to report.

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

### -Path
Workbook path.

```yaml
Type: String
Parameter Sets: Path
Aliases: InputPath, ReferencePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SkipCells
Skip visible cell value and formula comparison.

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

### -SkipCellStyles
Skip style-index comparison for used cells.

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

### -SkipComments
Skip legacy and threaded comment comparison.

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

### -SkipNamedRanges
Skip workbook and sheet-scoped named-range comparison.

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

### -SkipTables
Skip table metadata comparison.

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

### -SkipWorksheetMetadata
Skip worksheet view, validation, and filter metadata comparison.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `System.Management.Automation.PSObject`

## RELATED LINKS

- None

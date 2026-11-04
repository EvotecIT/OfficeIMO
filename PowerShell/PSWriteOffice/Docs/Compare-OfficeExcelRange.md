---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Compare-OfficeExcelRange
## SYNOPSIS
Compares two worksheets or ranges cell by cell.

## SYNTAX
### Path (Default)
```powershell
Compare-OfficeExcelRange [-ReferencePath] <string> [[-DifferencePath] <string>] [-Sheet <OfficeSheetSelector>] [-Range <string>] [-DifferenceSheet <OfficeSheetSelector>] [-DifferenceRange <string>] [-TrimStrings] [-IgnoreCase] [-StrictNullEmpty] [<CommonParameters>]
```

### Document
```powershell
Compare-OfficeExcelRange -ReferenceDocument <ExcelDocument> [-Sheet <OfficeSheetSelector>] [-Range <string>] [-DifferenceSheet <OfficeSheetSelector>] [-DifferenceRange <string>] [-TrimStrings] [-IgnoreCase] [-StrictNullEmpty] [-DifferenceDocument <ExcelDocument>] [<CommonParameters>]
```

## DESCRIPTION
Compares two worksheets or ranges cell by cell.

## EXAMPLES

### EXAMPLE 1
```powershell
Compare-OfficeExcelRange -ReferencePath .\Finance\Budget-v1.xlsx -DifferencePath .\Finance\Budget-v2.xlsx -Sheet Budget |
                Select-Object Kind, ReferenceCell, ReferenceValue, DifferenceValue
```

Lists the cells of the Budget sheet that differ between two versions of the workbook.

### EXAMPLE 2
```powershell
$compare = @{
    ReferencePath   = '.\Finance\Reconciliation-October.xlsx'
    Sheet           = 'Ledger'
    Range           = 'A1:F400'
    DifferenceSheet = 'Bank'
    DifferenceRange = 'A1:F400'
}
Compare-OfficeExcelRange @compare -TrimStrings -IgnoreCase |
    Export-OfficeExcel -Path .\Finance\Reconciliation-Differences.xlsx -Title 'Ledger and bank differences'
```

Compares the ledger with the bank statement range by range, ignoring case and stray spaces, and writes the differences to a review workbook.

## PARAMETERS

### -DifferenceDocument
Workbook document to compare against.

```yaml
Type: ExcelDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
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

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferenceRange
A1 range containing the comparison values in the difference workbook.

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

### -DifferenceSheet
Worksheet selector in the difference workbook.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IgnoreCase
Compare strings case-insensitively.

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

### -Range
A1 cell range on the selected worksheet.

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

### -ReferenceDocument
Caller-owned workbook containing the reference values.

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

### -ReferencePath
Path to the workbook containing the reference values.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrictNullEmpty
Treat null and empty string values as different.

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

### -TrimStrings
Compare strings after trimming whitespace.

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

- `None`

## RELATED LINKS

- None

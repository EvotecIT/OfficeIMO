---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Repair-OfficeExcelWorkbook
## SYNOPSIS
Runs OfficeIMO safe workbook repairs for common package, table, view, print, drawing, and calculation artifacts.

## SYNTAX
### Path (Default)
```powershell
Repair-OfficeExcelWorkbook [-Path] <string> [-SkipDefinedNames] [-SkipTables] [-SkipSheetViews] [-SkipPrintSettings] [-SkipDrawings] [-SkipCalculation] [-NoSave] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Repair-OfficeExcelWorkbook -Document <ExcelDocument> [-SkipDefinedNames] [-SkipTables] [-SkipSheetViews] [-SkipPrintSettings] [-SkipDrawings] [-SkipCalculation] [-NoSave] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Runs OfficeIMO safe workbook repairs for common package, table, view, print, drawing, and calculation artifacts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $repair = Repair-OfficeExcelWorkbook -Path .\QuarterlyReport.xlsx -PassThru
$repair.Actions | Format-Table Category,SheetName,Message
if ($repair.After.HasErrors) {
    $repair.After.Issues | Format-Table Severity,Category,SheetName,Address,Message
}
```

Uses the reusable OfficeIMO repair pipeline. The command normalizes safe workbook artifacts and returns before/after diagnostics when -PassThru is used.

## PARAMETERS

### -Document
Open workbook document to repair.

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

### -NoSave
Do not save after applying repairs to an open document.

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

### -PassThru
Emit a repair report object.

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
Workbook path to repair.

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

### -SkipCalculation
Skip calculation-chain cleanup and recalc-on-open metadata.

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

### -SkipDefinedNames
Skip defined-name repairs.

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

### -SkipDrawings
Skip drawing, image, and header/footer picture repairs.

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

### -SkipPrintSettings
Skip print, page-break, and page-scale repairs.

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

### -SkipSheetViews
Skip worksheet view and freeze-pane repairs.

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
Skip worksheet table repairs.

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

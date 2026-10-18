---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Test-OfficeExcelWorkbook
## SYNOPSIS
Runs OfficeIMO workbook diagnostics and optional safe repairs.

## SYNTAX
### Path (Default)
```powershell
Test-OfficeExcelWorkbook [-Path] <string> [-RepairDefinedNames] [-SkipOpenXmlValidation] [-Quiet] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Test-OfficeExcelWorkbook -Document <ExcelDocument> [-RepairDefinedNames] [-SkipOpenXmlValidation] [-Quiet] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Runs OfficeIMO workbook diagnostics and optional safe repairs.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doctor = Test-OfficeExcelWorkbook -Path .\Report.xlsx -SkipOpenXmlValidation
if (-not $doctor.Passed) {
    $doctor.Issues |
        Sort-Object Severity,Category,SheetName,Address |
        Format-Table Severity,Category,SheetName,Address,Message,RepairAction
}
```

Returns OfficeIMO workbook diagnostics for defined names, formulas, tables, drawings, connections, and package validation.

## PARAMETERS

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

### -Path
Workbook path.

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

### -Quiet
Return only a Boolean pass/fail value.

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

### -RepairDefinedNames
Repair duplicate, invalid, or broken defined names before reporting.

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

### -SkipOpenXmlValidation
Skip Open XML validator checks.

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

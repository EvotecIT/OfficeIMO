---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelTestOptions
## SYNOPSIS
Creates Excel checks for Test-OfficeDocument: capability preflight and package validation.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelTestOptions [-Capability <ExcelPreflightCapability[]>] [-NoOpenXmlValidation] [-IncludeFeatures] [-IncludeRepairHints] [<CommonParameters>]
```

## DESCRIPTION
Creates Excel checks for Test-OfficeDocument: capability preflight and package validation.

## EXAMPLES

### EXAMPLE 1
```powershell
Test-OfficeDocument -Path .\Finance\Budget-2027.xlsx -Options (New-OfficeExcelTestOptions -Capability CalculateFormulas)
```

Tests whether OfficeIMO can calculate the formulas in a workbook.

### EXAMPLE 2
```powershell
$checks = New-OfficeExcelTestOptions -Capability BindTemplate, ExportPdfReport -IncludeRepairHints
Get-ChildItem -Path .\Templates -Filter *.xltx |
    Test-OfficeDocument -Options $checks
```

Checks each template for template binding and PDF export and includes repair hints for anything that fails.

## PARAMETERS

### -Capability
Capabilities to evaluate or enforce. Defaults to all capabilities.

```yaml
Type: ExcelPreflightCapability[]
Parameter Sets: Default
Aliases: None
Possible values: ReadWorkbookData, EditCellValues, EditWorkbookStructure, UseCachedFormulaValues, CalculateFormulas, BindTemplate, ExportPdfReport

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeFeatures
Include discovered feature rows in object output.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeRepairHints
Include actionable OfficeIMO repair hints for blocked capabilities.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoOpenXmlValidation
Skips native Open XML schema validation while retaining other workbook checks.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None

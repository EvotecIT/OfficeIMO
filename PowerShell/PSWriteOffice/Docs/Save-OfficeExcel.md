---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Save-OfficeExcel
## SYNOPSIS
Saves an Excel workbook without disposing it.

## SYNTAX
### __AllParameterSets
```powershell
Save-OfficeExcel [-Document] <ExcelDocument> [-Path <string>] [-Open] [-Password <string>] [-SafePreflight] [-SafeRepairDefinedNames] [-ValidateOpenXml] [-DisableFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-ForceFullCalculationOnOpen] [-DateSystem <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Saves an Excel workbook without disposing it.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $workbook | Save-OfficeExcel
```

Writes pending changes to disk and keeps the workbook open.

## PARAMETERS

### -ClearCachedFormulaResults
Remove cached formula results before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DateSystem
Workbook date system for Excel date serials.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: 1900, 1904, NineteenHundred, NineteenFour

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DisableFastPackageWriter
Disable OfficeIMO fast package writers for this save.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Workbook to save.

```yaml
Type: ExcelDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -EvaluateFormulas
Evaluate supported formulas and write cached values before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ForceFullCalculationOnOpen
Request a full workbook recalculation when opened in Excel-compatible applications.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MarkFormulasDirty
Mark formula cells dirty before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the workbook after saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the workbook for further processing.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to save the workbook as an encrypted package.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional save-as path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SafePreflight
Run OfficeIMO worksheet preflight cleanup before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SafeRepairDefinedNames
Repair common defined-name issues before saving.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidateOpenXml
Validate the saved package with OpenXmlValidator and throw on errors.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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

- `OfficeIMO.Excel.ExcelDocument`

## RELATED LINKS

- None

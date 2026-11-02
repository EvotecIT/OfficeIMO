---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeExcelSaveOptions
## SYNOPSIS
Creates native workbook save settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficeExcelSaveOptions [-Preflight] [-RepairDefinedNames] [-ValidateOpenXml] [-NoFastPackageWriter] [-EvaluateFormulas] [-ClearCachedFormulaResults] [-MarkFormulasDirty] [-FullCalculationOnOpen] [<CommonParameters>]
```

## DESCRIPTION
Creates native workbook save settings.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeExcelSaveOptions -Preflight -EvaluateFormulas -FullCalculationOnOpen -ValidateOpenXml
Save-OfficeExcel -Document $workbook -SaveOptions $options
```


## PARAMETERS

### -ClearCachedFormulaResults
Remove cached formula results unless evaluation is requested.

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

### -EvaluateFormulas
Evaluate supported formulas and write their cached results.

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

### -FullCalculationOnOpen
Request a full recalculation when the workbook is opened.

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

### -MarkFormulasDirty
Mark formulas dirty for the next application open.

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

### -NoFastPackageWriter
Use the standard package finalization path.

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

### -Preflight
Run native worksheet preflight cleanup before saving.

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

### -RepairDefinedNames
Repair common defined-name issues before saving.

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

### -ValidateOpenXml
Validate the saved package against the Open XML schema.

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

- `OfficeIMO.Excel.ExcelSaveOptions`

## RELATED LINKS

- None

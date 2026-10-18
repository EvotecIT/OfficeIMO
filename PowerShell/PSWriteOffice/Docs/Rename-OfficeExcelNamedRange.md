---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Rename-OfficeExcelNamedRange
## SYNOPSIS
Renames a workbook or sheet-scoped Excel named range.

## SYNTAX
### Context (Default)
```powershell
Rename-OfficeExcelNamedRange [-Name] <string> [-NewName] <string> [-Global] [-ValidationMode <ExcelDefinedNameValidationMode>] [-PassThru] [-Save] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Rename-OfficeExcelNamedRange [-Name] <string> [-NewName] <string> -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-ValidationMode <ExcelDefinedNameValidationMode>] [-PassThru] [-Save] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Renames a workbook or sheet-scoped Excel named range.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $workbook = Get-OfficeExcel -Path .\Report.xlsx
$renamed = $workbook | Rename-OfficeExcelNamedRange -Name RevenueRange -NewName Revenue_Current -PassThru
Save-OfficeExcel -Document $workbook
```

Renames the defined name through OfficeIMO validation before saving the workbook.

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

### -Global
Use workbook-global scope from inside the DSL.

```yaml
Type: SwitchParameter
Parameter Sets: Context
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Name
Named range name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NewName
New named range name.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit a result object.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Save
Save the workbook immediately after renaming the name.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name for a sheet-scoped operation.

```yaml
Type: String
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Zero-based worksheet index for a sheet-scoped operation.

```yaml
Type: Int32
Parameter Sets: Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidationMode
Defined-name validation mode.

```yaml
Type: ExcelDefinedNameValidationMode
Parameter Sets: Context, Document
Aliases: None
Possible values: Sanitize, Strict

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

- `System.Boolean`

## RELATED LINKS

- None

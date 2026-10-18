---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Unprotect-OfficeExcelWorkbook
## SYNOPSIS
Removes workbook structure/window protection metadata.

## SYNTAX
### Context (Default)
```powershell
Unprotect-OfficeExcelWorkbook [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Unprotect-OfficeExcelWorkbook [-Path] <string> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Unprotect-OfficeExcelWorkbook -Document <ExcelDocument> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Removes workbook structure/window protection metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Unprotect-OfficeExcelWorkbook -Path .\Report.xlsx
            Test-OfficeExcelWorkbook -Path .\Report.xlsx -SkipOpenXmlValidation |
                Select-Object Passed, ProtectionSummary
```

Removes workbook structure/window protection metadata and saves the workbook.

## PARAMETERS

### -Document
Workbook to update outside the DSL context.

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

### -PassThru
Emit the workbook after removing protection.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Workbook path to update.

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

- `OfficeIMO.Excel.ExcelDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None

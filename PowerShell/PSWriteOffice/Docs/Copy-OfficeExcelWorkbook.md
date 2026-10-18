---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Copy-OfficeExcelWorkbook
## SYNOPSIS
Copies a workbook package while preserving package parts.

## SYNTAX
### __AllParameterSets
```powershell
Copy-OfficeExcelWorkbook [-Path] <string> [-DestinationPath] <string> [-Force] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Copies a workbook package while preserving package parts.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $copy = Copy-OfficeExcelWorkbook -Path .\Template.xlsx -DestinationPath .\Report.xlsx -Force -PassThru
Test-OfficeExcelWorkbook -Path $copy.FullName -SkipOpenXmlValidation |
    Select-Object Passed, WorksheetCount
```

Copies the workbook package and normalizes the workbook content type for the destination extension.

## PARAMETERS

### -DestinationPath
Destination workbook path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: Destination, OutputPath, TargetPath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Replace an existing destination workbook.

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

### -PassThru
Emit a FileInfo for the copied workbook.

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

### -Path
Source workbook or template package path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath, InputPath, SourcePath
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

- `None`

## OUTPUTS

- `System.IO.FileInfo`

## RELATED LINKS

- None

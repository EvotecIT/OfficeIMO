---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeExcelWorkbook
## SYNOPSIS
Converts Excel workbooks between supported .xls and .xlsx formats.

## SYNTAX
### __AllParameterSets
```powershell
ConvertTo-OfficeExcelWorkbook [-Path] <string> [-OutputPath] <string> [-Force] [-AllowLossyLegacyConversion] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Uses the OfficeIMO Excel normal load/save conversion path, including legacy XLS diagnostics and save preflight.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ConvertTo-OfficeExcelWorkbook -Path .\legacy.xls -OutputPath .\converted.xlsx -PassThru
```

Reads the .xls file and writes a .xlsx workbook.

### EXAMPLE 2
```powershell
PS> ConvertTo-OfficeExcelWorkbook -Path .\report.xlsx -OutputPath .\report.xls -Force
```

Writes a supported native BIFF8 .xls workbook.

## PARAMETERS

### -AllowLossyLegacyConversion
Allow conversion when a legacy XLS source contains unsupported or preserve-only content.

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

### -Force
Overwrite an existing destination file.

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
Open the converted workbook after saving.

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

### -OutputPath
Destination .xls or .xlsx file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: OutPath
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the saved file information.

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
Source .xls or .xlsx file path.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: FilePath
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

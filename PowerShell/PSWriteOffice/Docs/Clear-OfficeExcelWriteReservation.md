---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Clear-OfficeExcelWriteReservation
## SYNOPSIS
Clears workbook write-reservation metadata.

## SYNTAX
### Context (Default)
```powershell
Clear-OfficeExcelWriteReservation [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Clear-OfficeExcelWriteReservation [-Path] <string> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Clear-OfficeExcelWriteReservation -Document <ExcelDocument> [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Clears workbook write-reservation metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Clear-OfficeExcelWriteReservation -Path .\Report.xlsx
```

Removes the workbook file-sharing/write-reservation node while leaving workbook protection and encryption state unchanged.

## PARAMETERS

### -Document
Open workbook document to update.

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
Emit the resulting write-reservation metadata.

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
Aliases: FilePath, InputPath, FullName
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

- `None`

## RELATED LINKS

- None

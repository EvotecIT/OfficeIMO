---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Clear-OfficeExcelPageBreak
## SYNOPSIS
Clears manual row or column page breaks from an Excel worksheet.

## SYNTAX
### Context (Default)
```powershell
Clear-OfficeExcelPageBreak [-Sheet <string>] [-SheetIndex <Int32>] [-Row <int[]>] [-Column <int[]>] [-All] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Clear-OfficeExcelPageBreak [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <int[]>] [-Column <int[]>] [-All] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Clear-OfficeExcelPageBreak -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <int[]>] [-Column <int[]>] [-All] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Clears manual row or column page breaks from an Excel worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Clear-OfficeExcelPageBreak -Path .\Report.xlsx -Sheet Data -Row 25 -Column 8 -Confirm:$false
            Get-OfficeExcelPageBreak -Path .\Report.xlsx -Sheet Data |
                Format-Table Type, Position, SheetName
```

Removes the selected row and column page breaks and saves the workbook.

## PARAMETERS

### -All
Clear all manual row and column page breaks from selected worksheets.

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

### -Column
One-based columns whose manual page breaks should be removed.

```yaml
Type: Int32[]
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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
Emit the object created or changed by the command.

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

### -Row
One-based rows whose manual page breaks should be removed.

```yaml
Type: Int32[]
Parameter Sets: Context, Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name to update. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: String
Parameter Sets: Context, Path, Document
Aliases: WorksheetName
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SheetIndex
Worksheet index (0-based) to update. Defaults to the current DSL sheet or all workbook sheets.

```yaml
Type: Int32
Parameter Sets: Context, Path, Document
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

- `None`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelPageBreak
## SYNOPSIS
Adds manual row or column page breaks to an Excel worksheet.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeExcelPageBreak [-Sheet <string>] [-SheetIndex <Int32>] [-Row <int[]>] [-Column <int[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Add-OfficeExcelPageBreak [-Path] <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <int[]>] [-Column <int[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeExcelPageBreak -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <int[]>] [-Column <int[]>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Adds manual row or column page breaks to an Excel worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Add-OfficeExcelPageBreak -Row 25,50 -Column 8 }
```

Adds row page breaks after rows 25 and 50 and a column page break after column 8.

## PARAMETERS

### -Column
One-based columns after which manual page breaks should be inserted.

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
Emit page-break records after adding them.

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
One-based rows after which manual page breaks should be inserted.

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
Worksheet name. Defaults to the current sheet inside an ExcelSheet block.

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
Worksheet index when using a workbook object or path.

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

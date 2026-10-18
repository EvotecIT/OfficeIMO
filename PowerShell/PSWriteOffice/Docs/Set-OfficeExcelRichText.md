---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelRichText
## SYNOPSIS
Sets mixed-format rich text runs in an Excel cell.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelRichText -Run <Object[]> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Path
```powershell
Set-OfficeExcelRichText [-Path] <string> -Run <Object[]> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelRichText -Document <ExcelDocument> -Run <Object[]> [-Sheet <string>] [-SheetIndex <Int32>] [-Row <Int32>] [-Column <Int32>] [-Address <string>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Sets mixed-format rich text runs in an Excel cell.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Summary' {
    Set-OfficeExcelRichText -Address A1 -Run 'Status: ', @{ Text = 'Blocked'; Bold = $true; Color = '#C00000' }
}
```

Stores inline rich text in the target cell using OfficeIMO's reusable rich-text cell model.

## PARAMETERS

### -Address
A1-style cell address.

```yaml
Type: String
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
1-based column index.

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
Emit written rich text runs.

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
1-based row index.

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

### -Run
Rich text runs. Each run can be a string, hashtable, PSCustomObject, or ExcelRichTextRun.

```yaml
Type: Object[]
Parameter Sets: Context, Path, Document
Aliases: Runs
Possible values:

Required: True
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

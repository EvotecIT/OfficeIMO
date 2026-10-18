---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Update-OfficeExcelText
## SYNOPSIS
Replaces text in worksheet values.

## SYNTAX
### Path (Default)
```powershell
Update-OfficeExcelText [-Path] <string> -OldValue <string> -NewValue <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-CaseSensitive] [-Regex] [-Open] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Update-OfficeExcelText -Document <ExcelDocument> -OldValue <string> -NewValue <string> [-Sheet <string>] [-SheetIndex <Int32>] [-Range <string>] [-CaseSensitive] [-Regex] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Replaces text in worksheet values.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $count = Update-OfficeExcelText -Path .\Report.xlsx -Sheet Summary -OldValue Draft -NewValue Ready -PassThru
[pscustomobject]@{
    Path = '.\Report.xlsx'
    Replacements = $count
}
```

Updates matching text cells on a sheet, saves the workbook, and returns the replacement count.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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

### -NewValue
Replacement text.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OldValue
Text or pattern to replace.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the file after saving when using -Path.

```yaml
Type: SwitchParameter
Parameter Sets: Path
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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

### -Range
A1 range to update. Defaults to each selected worksheet's used range.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Regex
Treat -OldValue as a regular expression.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name. Defaults to all sheets for path/document use and current sheet inside an ExcelSheet block.

```yaml
Type: String
Parameter Sets: Path, Document
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
Parameter Sets: Path, Document
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

- `System.Int32`

## RELATED LINKS

- None

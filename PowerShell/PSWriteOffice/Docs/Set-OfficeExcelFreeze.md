---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeExcelFreeze
## SYNOPSIS
Freezes panes on the current worksheet.

## SYNTAX
### Context (Default)
```powershell
Set-OfficeExcelFreeze [-TopRows <int>] [-LeftColumns <int>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Set-OfficeExcelFreeze -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-TopRows <int>] [-LeftColumns <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Freezes panes on the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelFreeze -TopRows 1 }
```

Freezes the first row.

### EXAMPLE 2
```powershell
PS> ExcelSheet 'Data' { Set-OfficeExcelFreeze -TopRows 1 -LeftColumns 1 }
```

Freezes row 1 and column A.

## PARAMETERS

### -Document
Workbook to operate on outside the DSL context.

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

### -LeftColumns
Number of columns to freeze from the left.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
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
Worksheet name when using Document.

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
Worksheet index (0-based) when using Document.

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

### -TopRows
Number of rows to freeze from the top.

```yaml
Type: Int32
Parameter Sets: Context, Document
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

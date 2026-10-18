---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Protect-OfficeExcelSheet
## SYNOPSIS
Protects the current worksheet.

## SYNTAX
### Context (Default)
```powershell
Protect-OfficeExcelSheet [-AllowTableEditing] [-AllowSelectLockedCells <bool>] [-AllowSelectUnlockedCells <bool>] [-AllowFormatCells <bool>] [-AllowFormatColumns <bool>] [-AllowFormatRows <bool>] [-AllowInsertColumns <bool>] [-AllowInsertRows <bool>] [-AllowInsertHyperlinks <bool>] [-AllowDeleteColumns <bool>] [-AllowDeleteRows <bool>] [-AllowSort <bool>] [-AllowAutoFilter <bool>] [-AllowPivotTables <bool>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Protect-OfficeExcelSheet -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-AllowTableEditing] [-AllowSelectLockedCells <bool>] [-AllowSelectUnlockedCells <bool>] [-AllowFormatCells <bool>] [-AllowFormatColumns <bool>] [-AllowFormatRows <bool>] [-AllowInsertColumns <bool>] [-AllowInsertRows <bool>] [-AllowInsertHyperlinks <bool>] [-AllowDeleteColumns <bool>] [-AllowDeleteRows <bool>] [-AllowSort <bool>] [-AllowAutoFilter <bool>] [-AllowPivotTables <bool>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Protects the current worksheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Protect-OfficeExcelSheet }
```

Enables worksheet protection.

## PARAMETERS

### -AllowAutoFilter
Allow AutoFilter.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowDeleteColumns
Allow deleting columns.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowDeleteRows
Allow deleting rows.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowFormatCells
Allow formatting cells.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowFormatColumns
Allow formatting columns.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowFormatRows
Allow formatting rows.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowInsertColumns
Allow inserting columns.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowInsertHyperlinks
Allow inserting hyperlinks.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowInsertRows
Allow inserting rows.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowPivotTables
Allow PivotTables.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowSelectLockedCells
Allow selecting locked cells.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowSelectUnlockedCells
Allow selecting unlocked cells.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowSort
Allow sorting.

```yaml
Type: Boolean
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowTableEditing
Allow common protected-table workflows: selecting cells, inserting rows, sorting, and filtering.

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

### -PassThru
Emit the worksheet after protection.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None

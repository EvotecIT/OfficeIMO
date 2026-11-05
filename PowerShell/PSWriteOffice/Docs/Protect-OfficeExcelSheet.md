---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Protect-OfficeExcelSheet
## SYNOPSIS
Protects a worksheet and chooses what users may still do.

## SYNTAX
### Default (Default)
```powershell
Protect-OfficeExcelSheet [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-Password <securestring>] [-AllowTableEditing] [-NoSelectLockedCells] [-NoSelectUnlockedCells] [-AllowFormatCells] [-AllowFormatColumns] [-AllowFormatRows] [-AllowInsertColumns] [-AllowInsertRows] [-AllowInsertHyperlinks] [-AllowDeleteColumns] [-AllowDeleteRows] [-AllowSort] [-AllowAutoFilter] [-AllowPivotTables] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Protects a worksheet and chooses what users may still do.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
    Protect-OfficeExcelSheet -Sheet Inputs
}
```

Locks the Inputs sheet so only unlocked cells can be edited.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\HR\Staff-Directory.xlsx {
    ExcelSheet 'Directory' {
        Protect-OfficeExcelSheet -AllowSort -AllowAutoFilter -AllowFormatColumns
    }
}
```

Protects the staff list while still letting colleagues sort, filter and resize columns.

### EXAMPLE 3
```powershell
$password = Read-Host -Prompt 'Sheet password' -AsSecureString
Open-OfficeExcel -Path .\Finance\Pricing-Model.xlsx {
    Protect-OfficeExcelSheet -Sheet Inputs -Password $password -AllowSort -AllowAutoFilter
}
```

Locks the input sheet with a password read from the console, still allowing sorting and filtering.

## PARAMETERS

### -AllowAutoFilter
Allow AutoFilter.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Type: SwitchParameter
Parameter Sets: Default
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
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned workbook; omission uses the enclosing Excel content block.

```yaml
Type: ExcelDocument
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -NoSelectLockedCells
Prevents selecting locked cells while the worksheet is protected.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSelectUnlockedCells
Prevents selecting unlocked cells while the worksheet is protected.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emits the native result after the operation.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used to open an encrypted workbook package.

```yaml
Type: SecureString
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Worksheet name, one-based position or native worksheet object. Multi-sheet commands also accept wildcard names.

```yaml
Type: OfficeSheetSelector
Parameter Sets: Default
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

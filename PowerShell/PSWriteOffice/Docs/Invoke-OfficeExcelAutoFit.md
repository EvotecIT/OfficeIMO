---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Invoke-OfficeExcelAutoFit
## SYNOPSIS
Automatically fits Excel row heights and/or column widths.

## SYNTAX
### Context (Default)
```powershell
Invoke-OfficeExcelAutoFit [-Columns] [-Rows] [-All] [-Column <int[]>] [-Row <int[]>] [<CommonParameters>]
```

### Document
```powershell
Invoke-OfficeExcelAutoFit -Document <ExcelDocument> [-Sheet <string>] [-SheetIndex <Int32>] [-Columns] [-Rows] [-All] [-Column <int[]>] [-Row <int[]>] [<CommonParameters>]
```

## DESCRIPTION
Targets the current worksheet context or a worksheet on a supplied document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> ExcelSheet 'Data' { Invoke-OfficeExcelAutoFit -Columns }
```

Adjusts column widths for the active sheet.

## PARAMETERS

### -All
Auto-fit both rows and columns.

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

### -Column
Auto-fit specific column indexes (1-based).

```yaml
Type: Int32[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Columns
Auto-fit all columns.

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

### -Row
Auto-fit specific row indexes (1-based).

```yaml
Type: Int32[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rows
Auto-fit all rows.

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

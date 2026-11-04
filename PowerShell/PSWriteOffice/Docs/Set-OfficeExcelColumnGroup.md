---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelColumnGroup
## SYNOPSIS
Groups columns into a collapsible outline, or clears grouping.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelColumnGroup [-StartColumn] <Object> [[-EndColumn] <Object>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-OutlineLevel <int>] [-Collapsed] [-Hidden] [-Clear] [-KeepHidden] [-SummaryLeft] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Groups columns into a collapsible outline, or clears grouping.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
    ExcelSheet 'Departments' {
        ExcelColumnGroup B M
    }
}
```

Groups columns B to M, the monthly figures, next to the yearly total.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Forecast.xlsx {
    ExcelSheet 'Forecast' {
        ExcelColumnGroup B D -Collapsed
        ExcelColumnGroup F H -Collapsed
        ExcelColumnGroup J L -Collapsed
        ExcelColumnGroup N P -Collapsed
    }
}
```

Groups each quarter's three months and collapses them, so the sheet shows the quarterly totals first.

## PARAMETERS

### -Clear
Clear column grouping metadata from the target range.

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

### -Collapsed
Hide the grouped columns and mark the following summary column as collapsed.

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

### -EndColumn
Last 1-based column in the group. Defaults to the start column.

```yaml
Type: Object
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
Hide the grouped columns without marking the group collapsed.

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

### -KeepHidden
Keep hidden columns hidden when clearing column grouping metadata.

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

### -OutlineLevel
Excel outline level from 1 through 7.

```yaml
Type: Int32
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

### -StartColumn
First 1-based column in the group.

```yaml
Type: Object
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SummaryLeft
Places outline summary columns to the left; explicit false places them to the right.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Excel.ExcelDocument`

## OUTPUTS

- `None`

## RELATED LINKS

- None

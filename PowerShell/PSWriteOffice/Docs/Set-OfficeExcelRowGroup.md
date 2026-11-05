---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelRowGroup
## SYNOPSIS
Groups rows into a collapsible outline, or clears grouping.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelRowGroup [-StartRow] <int> [[-EndRow] <Int32>] [-Document <ExcelDocument>] [-Sheet <OfficeSheetSelector>] [-OutlineLevel <int>] [-Collapsed] [-Hidden] [-Clear] [-KeepHidden] [-SummaryAbove] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Groups rows into a collapsible outline, or clears grouping.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Budget-2027.xlsx {
    ExcelSheet 'Departments' {
        ExcelRowGroup 3 12
    }
}
```

Groups rows 3 to 12 so they can be collapsed under their summary row.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Regional-Sales.xlsx {
    ExcelSheet 'Detail' {
        ExcelRowGroup 3 40 -Collapsed -SummaryAbove
        ExcelRowGroup 42 77 -Collapsed -SummaryAbove
        ExcelRowGroup 79 120 -Collapsed -SummaryAbove
    }
}
```

Groups and collapses the detail rows of each region with the summary row above them, so the sheet opens showing only the regional totals.

## PARAMETERS

### -Clear
Clear row grouping metadata from the target range.

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
Hide the grouped rows and mark the following summary row as collapsed.

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

### -EndRow
Last 1-based row in the group. Defaults to StartRow.

```yaml
Type: Int32
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
Hide the grouped rows without marking the group collapsed.

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
Keep hidden rows hidden when clearing row grouping metadata.

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

### -StartRow
First 1-based row in the group.

```yaml
Type: Int32
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SummaryAbove
Places outline summary rows above detail rows; explicit false places them below.

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

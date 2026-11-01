---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Remove-OfficeExcelPageBreak
## SYNOPSIS
Removes manual print page breaks from selected open worksheets.

## SYNTAX
### Default (Default)
```powershell
Remove-OfficeExcelPageBreak [-Sheet <OfficeSheetSelector[]>] [-Row <int[]>] [-Column <int[]>] [-All] [-PassThru] [-Document <ExcelDocument>] [<CommonParameters>]
```

## DESCRIPTION
Removes manual print page breaks from selected open worksheets.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path ./Report.xlsx { Remove-OfficeExcelPageBreak -Sheet 'Report*' -All }
```


## PARAMETERS

### -All
Remove every manual row and column break on the selected worksheets.

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

### -Column
One-based columns whose manual breaks are removed.

```yaml
Type: Int32[]
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
Gets or sets an explicit live document; omission uses the enclosing content block.

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

### -PassThru
Emit a snapshot for each break actually removed.

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

### -Row
One-based rows whose manual breaks are removed.

```yaml
Type: Int32[]
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
Worksheet names, wildcard patterns, one-based positions or native worksheets.

```yaml
Type: OfficeSheetSelector[]
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

- `PSWriteOffice.Models.Excel.ExcelPageBreakInfo`: A manual print page-break snapshot that remains readable after its source workbook is closed.

## RELATED LINKS

- None

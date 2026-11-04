---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelSlicer
## SYNOPSIS
Adds slicer or timeline binding metadata for a pivot table or table.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelSlicer -Name <string> [-Document <ExcelDocument>] [-SourceName <string>] [-PivotTable <string>] [-Xml <string>] [-Timeline] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds slicer or timeline binding metadata for a pivot table or table.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
    ExcelSlicer -Name RegionSlicer -SourceName Region -PivotTable SalesPivot
}
```

Adds slicer metadata for the Region field of the SalesPivot pivot table.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
    ExcelSlicer -Name OrderDateTimeline -SourceName OrderDate -PivotTable SalesPivot -Timeline
}
```

Adds timeline metadata for the OrderDate field so readers can filter the pivot table by period in Excel.

## PARAMETERS

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

### -Name
Slicer cache name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
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

### -PivotTable
Native pivot table whose data or entire range is targeted.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SourceName
Source field, table column, or pivot field name.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Timeline
Creates timeline metadata instead of a standard slicer.

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

### -Xml
Caller-supplied slicer cache XML. When provided, OfficeIMO preserves it within its pivot-interaction metadata envelope.

```yaml
Type: String
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

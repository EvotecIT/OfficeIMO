---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeExcelRefreshOnOpen
## SYNOPSIS
Turns refresh-on-open on or off for pivot caches and connections.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeExcelRefreshOnOpen [-Document <ExcelDocument>] [-PivotTables] [-Connections] [-Disable] [-SavePivotSourceData] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Turns refresh-on-open on or off for pivot caches and connections.

## EXAMPLES

### EXAMPLE 1
```powershell
Open-OfficeExcel -Path .\Reports\Sales-2026.xlsx {
    ExcelRefreshOnOpen -PivotTables
}
```

Makes Excel refresh every pivot cache when the workbook opens.

### EXAMPLE 2
```powershell
Open-OfficeExcel -Path .\Archive\Sales-2026-09.xlsx {
    ExcelRefreshOnOpen -PivotTables -Connections -Disable -PassThru
}
```

Turns off refresh-on-open for pivot tables and connections in a month-end copy, so the archived figures stay as they were.

## PARAMETERS

### -Connections
Update workbook connection refresh-on-open metadata.

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

### -Disable
Disable refresh-on-open instead of enabling it.

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

### -PivotTables
Update pivot cache refresh-on-open metadata.

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

### -SavePivotSourceData
Preserve pivot cache source data.

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

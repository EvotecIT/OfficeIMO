---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeExcelReportLegend
## SYNOPSIS
Adds a legend table that explains values or colors.

## SYNTAX
### Default (Default)
```powershell
Add-OfficeExcelReportLegend [[-Title] <string>] -Header <string[]> -InputObject <Object[]> [-FirstColumnFillByValue <hashtable>] [-HeaderFillColor <OfficeColor>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a legend table that explains values or colors.

## EXAMPLES

### EXAMPLE 1
```powershell
$legend = @(
    [pscustomobject]@{ Status = 'Ready'; Meaning = 'Validated and ready to ship' }
    [pscustomobject]@{ Status = 'Review'; Meaning = 'Waiting for the owner' }
)
New-OfficeExcel -Path .\Reports\Release-Readiness.xlsx {
    ExcelReportSheet 'Readiness' {
        ExcelReportLegend 'Status' -Header Status, Meaning -InputObject $legend
    }
}
```

Explains the status values used in the report.

### EXAMPLE 2
```powershell
$legend = @(
    [pscustomobject]@{ Status = 'Ready'; Meaning = 'Validated and ready to ship' }
    [pscustomobject]@{ Status = 'Review'; Meaning = 'Waiting for the owner' }
    [pscustomobject]@{ Status = 'Blocked'; Meaning = 'Cannot ship until a dependency is fixed' }
)
New-OfficeExcel -Path .\Reports\Release-Readiness.xlsx {
    ExcelReportSheet 'Readiness' {
        ExcelReportTable (Import-OfficeCsv -Path .\Data\Release-Items.csv) 'Release 2026.11'
        ExcelReportLegend 'Status' -Header Status, Meaning -InputObject $legend -FirstColumnFillByValue @{ Ready = '#DCFCE7'; Review = '#FEF3C7'; Blocked = '#FEE2E2' }
    }
}
```

Colours each status in the legend the same way as the rows of the readiness table.

## PARAMETERS

### -CaseSensitive
Matches names and text using ordinal case sensitivity. Omission ignores case.

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

### -FirstColumnFillByValue
Optional first-column fill colors keyed by first-column value.

```yaml
Type: Hashtable
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Header text identifying the target column.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HeaderFillColor
Optional header fill color.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Objects, dictionaries or native tabular data supplied directly or through the pipeline.

```yaml
Type: Object[]
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

### -Title
Chart title.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeExcelReportLegend
## SYNOPSIS
Adds a legend table to the current Excel report sheet.

## SYNTAX
### __AllParameterSets
```powershell
Add-OfficeExcelReportLegend [[-Title] <string>] -Header <string[]> -InputObject <Object[]> [-FirstColumnFillByValue <hashtable>] [-HeaderFillColor <string>] [-CaseSensitive] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a legend table to the current Excel report sheet.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $legendRows = @(
    [pscustomobject]@{ Status = 'Ready'; Meaning = 'Validated and ready' }
    [pscustomobject]@{ Status = 'Review'; Meaning = 'Needs owner review' }
)
New-OfficeExcel -Path .\Operations.xlsx {
    Add-OfficeExcelReportSheet -Name Summary {
        Add-OfficeExcelReportLegend -Title 'Status legend' -Header Status, Meaning -InputObject $legendRows -FirstColumnFillByValue @{ Ready = '#d9f7be'; Review = '#fff7e6' }
    }
}
```

Renders legend rows and applies optional fill colors keyed by the first column.

## PARAMETERS

### -CaseSensitive
Use case-sensitive matching for first-column fill values.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Header
Column headers.

```yaml
Type: String[]
Parameter Sets: __AllParameterSets
Aliases: Headers
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
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
Rows. Each row may be an array, enumerable, hashtable, or object.

```yaml
Type: Object[]
Parameter Sets: __AllParameterSets
Aliases: Rows
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the object created or changed by the command.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional legend title.

```yaml
Type: String
Parameter Sets: __AllParameterSets
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

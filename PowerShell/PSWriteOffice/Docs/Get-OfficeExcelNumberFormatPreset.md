---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelNumberFormatPreset
## SYNOPSIS
Lists OfficeIMO Excel number format presets and their format codes.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeExcelNumberFormatPreset [-Decimals <int>] [-CultureName <string>] [<CommonParameters>]
```

## DESCRIPTION
Lists OfficeIMO Excel number format presets and their format codes.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeExcelNumberFormatPreset -CultureName en-US -Decimals 2 |
                Where-Object Name -eq Currency
```

Returns the preset name and the Excel number format code that PSWriteOffice cmdlets can apply to cells or columns.

## PARAMETERS

### -CultureName
Culture name used for currency symbols, such as en-US or pl-PL.

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

### -Decimals
Decimal places used for decimal, percent, currency, and scientific presets.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None

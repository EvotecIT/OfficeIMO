---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficeExcelTableStyle
## SYNOPSIS
Lists built-in Excel table styles and compatibility recommendations.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficeExcelTableStyle [-Profile <ExcelTableStyleCompatibilityProfile>] [-RecommendedOnly] [<CommonParameters>]
```

## DESCRIPTION
Lists built-in Excel table styles and compatibility recommendations.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficeExcelTableStyle -RecommendedOnly |
                Sort-Object Name |
                Format-Table Name, Profile
```

Uses OfficeIMO's table style catalog to return styles that are broadly stable across desktop, web, and spreadsheet viewers.

## PARAMETERS

### -Profile
Compatibility profile used to evaluate table styles.

```yaml
Type: ExcelTableStyleCompatibilityProfile
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Desktop, CrossHost

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RecommendedOnly
Return only styles recommended for the selected profile.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None

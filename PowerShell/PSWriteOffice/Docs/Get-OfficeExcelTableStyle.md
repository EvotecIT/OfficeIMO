---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficeExcelTableStyle
## SYNOPSIS
Lists built-in Excel table styles and compatibility recommendations.

## SYNTAX
### Default (Default)
```powershell
Get-OfficeExcelTableStyle [-Profile <ExcelTableStyleCompatibilityProfile>] [-RecommendedOnly] [<CommonParameters>]
```

## DESCRIPTION
Lists built-in Excel table styles and compatibility recommendations.

## EXAMPLES

### EXAMPLE 1
```powershell
Get-OfficeExcelTableStyle -RecommendedOnly
```

Shows the built-in table styles that render well in every supported host.

### EXAMPLE 2
```powershell
Get-OfficeExcelTableStyle -Profile CrossHost -RecommendedOnly | Format-Table
```

Lists the styles recommended for workbooks also opened in web and mobile spreadsheet apps.

## PARAMETERS

### -Profile
Compatibility profile used to evaluate table styles.

```yaml
Type: ExcelTableStyleCompatibilityProfile
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `None`

## RELATED LINKS

- None

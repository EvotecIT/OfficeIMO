---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointSection
## SYNOPSIS
Reads sections from a live presentation, optionally matching a PowerShell wildcard pattern.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointSection [[-Document] <PowerPointDocument>] [-Name <string>] [-CaseSensitive] [<CommonParameters>]
```

## DESCRIPTION
Reads sections from a live presentation, optionally matching a PowerShell wildcard pattern.

## EXAMPLES

### EXAMPLE 1
```powershell
$deck = Open-OfficePowerPoint -Path ./Review.pptx
Get-OfficePowerPointSection -Document $deck
Close-OfficePowerPoint -Document $deck
```


### EXAMPLE 2
```powershell
$deck | Get-OfficePowerPointSection -Name Appendix
```


## PARAMETERS

### -CaseSensitive
Gets or sets whether name matching is case-sensitive.

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

### -Document
Gets or sets the live document; omission uses its enclosing content block.

```yaml
Type: PowerPointDocument
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Name
Gets or sets a section name or PowerShell wildcard pattern.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: True
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSectionInfo`

## RELATED LINKS

- None

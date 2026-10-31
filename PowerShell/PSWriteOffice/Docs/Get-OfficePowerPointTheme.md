---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Get-OfficePowerPointTheme
## SYNOPSIS
Reads the selected slide master's theme, colors and fonts.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointTheme [[-Document] <PowerPointDocument>] [-Master <int>] [<CommonParameters>]
```

## DESCRIPTION
Reads the selected slide master's theme, colors and fonts.

## EXAMPLES

### EXAMPLE 1
```powershell
$deck = Open-OfficePowerPoint -Path ./Review.pptx
Get-OfficePowerPointTheme -Document $deck -Master 1
Close-OfficePowerPoint -Document $deck
```


### EXAMPLE 2
```powershell
New-OfficePowerPoint -Path ./Review.pptx -Theme Modern {
    $theme = Get-OfficePowerPointTheme
    PowerPointSlide -Title 'Service review'
}
```


## PARAMETERS

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

### -Master
Gets or sets a one-based master position; defaults to the first.

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

- `OfficeIMO.PowerPoint.PowerPointDocument`

## OUTPUTS

- `PSWriteOffice.Services.PowerPoint.PowerPointThemeInfo`: Describes a PowerPoint theme for a single slide master.

## RELATED LINKS

- None

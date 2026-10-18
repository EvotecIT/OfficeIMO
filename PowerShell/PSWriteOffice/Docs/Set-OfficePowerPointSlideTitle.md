---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideTitle
## SYNOPSIS
Sets the text of the title placeholder on a slide.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointSlideTitle -Title <string> [-Slide <PowerPointSlide>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Targets the title placeholder when available; otherwise adds a new title shape.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTitle -Title 'Executive Summary'
```

Updates the first slide’s title.

## PARAMETERS

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

### -Slide
Slide whose title should change (optional inside DSL).

```yaml
Type: PowerPointSlide
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Title
New title text.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `None`

## RELATED LINKS

- None

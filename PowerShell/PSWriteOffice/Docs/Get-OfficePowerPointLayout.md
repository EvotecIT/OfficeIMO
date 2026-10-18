---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointLayout
## SYNOPSIS
Lists slide layouts available in a presentation.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointLayout [-Presentation <PowerPointPresentation>] [-Master <int>] [<CommonParameters>]
```

## DESCRIPTION
Lists slide layouts available in a presentation.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointLayout -Presentation $ppt
```

Returns layout metadata including name, type, and index.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx { $layout = @(Get-OfficePowerPointLayout)[0] }
```

Uses the current DSL presentation context.

## PARAMETERS

### -Master
Slide master index.

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

### -Presentation
Presentation to inspect (optional inside DSL).

```yaml
Type: PowerPointPresentation
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlideLayoutInfo`

## RELATED LINKS

- None

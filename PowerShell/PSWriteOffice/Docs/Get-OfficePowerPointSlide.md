---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointSlide
## SYNOPSIS
Enumerates slides or retrieves a specific slide.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointSlide [-Presentation <PowerPointPresentation>] [-Index <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Supports pipeline-friendly iteration over Slides or direct index selection.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt | Get-OfficePowerPointPlaceholder -PlaceholderType Title | Select-Object -ExpandProperty Text
```

Streams each slide so you can read the title placeholder text.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx { Get-OfficePowerPointSlide | Select-Object -First 1 }
```

Uses the current DSL presentation context.

## PARAMETERS

### -Index
Optional zero-based index; omit to enumerate all slides.

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

- `None`

## RELATED LINKS

- None

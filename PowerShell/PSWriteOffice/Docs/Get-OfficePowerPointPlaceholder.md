---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Get-OfficePowerPointPlaceholder
## SYNOPSIS
Gets placeholder text boxes from a slide.

## SYNTAX
### __AllParameterSets
```powershell
Get-OfficePowerPointPlaceholder [-Slide <PowerPointSlide>] [-PlaceholderType <PowerPointPlaceholderType>] [-Index <UInt32>] [<CommonParameters>]
```

## DESCRIPTION
Gets placeholder text boxes from a slide.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointPlaceholder -Slide $slide -PlaceholderType Title
```

Returns the title placeholder textbox if present.

### EXAMPLE 2
```powershell
PS> New-OfficePowerPoint -Path .\deck.pptx { PptSlide { Get-OfficePowerPointPlaceholder -PlaceholderType Title } }
```

Uses the current slide context.

## PARAMETERS

### -Index
Optional placeholder index.

```yaml
Type: UInt32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PlaceholderType
Placeholder type to filter on.

```yaml
Type: PowerPointPlaceholderType
Parameter Sets: __AllParameterSets
Aliases: Type
Possible values: Title, Body, CenteredTitle, SubTitle, DateAndTime, SlideNumber, Footer, Header, Object, Chart, Table, ClipArt, Diagram, Media, SlideImage, Picture

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Slide to inspect (optional inside DSL).

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None

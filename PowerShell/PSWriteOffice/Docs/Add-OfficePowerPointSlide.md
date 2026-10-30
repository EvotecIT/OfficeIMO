---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointSlide
## SYNOPSIS
Adds a designed slide, choosing its layout from typed content unless Layout is supplied.

## SYNTAX
### Content (Default)
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] [-Title <string>] [-Subtitle <string>] [-Eyebrow <string>] [-MapLabel <string>] [-Layout <PowerPointLayout>] [-Seed <string>] [-NoFooter] [-NoSlideNumber] [-Background <OfficeColor>] [-BackgroundImage <string>] [-Transition <PowerPointSlideTransition>] [-PassThru] [-Document <PowerPointPresentation>] [<CommonParameters>]
```

### Template
```powershell
Add-OfficePowerPointSlide [[-Content] <scriptblock>] -TemplateLayout <PowerPointLayoutSelector> [-Title <string>] [-Subtitle <string>] [-Eyebrow <string>] [-MapLabel <string>] [-Layout <PowerPointLayout>] [-Master <int>] [-Seed <string>] [-NoFooter] [-NoSlideNumber] [-Background <OfficeColor>] [-BackgroundImage <string>] [-Transition <PowerPointSlideTransition>] [-PassThru] [-Document <PowerPointPresentation>] [<CommonParameters>]
```

## DESCRIPTION
Adds a designed slide, choosing its layout from typed content unless Layout is supplied.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePowerPoint -Path .\Review.pptx {
    PowerPointSlide -Title 'Service review' -Subtitle 'October 2026'
    PowerPointSlide -Title 'Signals' { PowerPointCard 'Availability' -Text 'All regions healthy' }
}
```


## PARAMETERS

### -Background
Gets or sets the background color for this slide and its continuations, including theme roles.

```yaml
Type: OfficeColor
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundImage
Gets or sets a literal local background image path for this slide and its continuations.

```yaml
Type: String
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets the slide's typed content block.

```yaml
Type: ScriptBlock
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets an explicit live document; omission uses the enclosing content block.

```yaml
Type: PowerPointPresentation
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Eyebrow
Gets or sets the small label above the title.

```yaml
Type: String
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Layout
Gets or sets a designed layout; omission selects it from content.

```yaml
Type: PowerPointLayout
Parameter Sets: Content, Template
Aliases: None
Possible values: Auto, Title, Section, TitleAndContent, TwoContent, Comparison, Kpi, Cards, Process, Quote, ImageFull, Closing, CaseStudy, Logos, Coverage

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MapLabel
Gets or sets the caption of the coverage panel.

```yaml
Type: String
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Master
Gets or sets a one-based master position for a template layout; defaults to the first.

```yaml
Type: Int32
Parameter Sets: Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoFooter
Gets or sets whether this slide omits its footer.

```yaml
Type: SwitchParameter
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSlideNumber
Gets or sets whether this slide omits its slide number.

```yaml
Type: SwitchParameter
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether the created physical slides are returned.

```yaml
Type: SwitchParameter
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Seed
Gets or sets a deterministic design seed for this slide.

```yaml
Type: String
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Subtitle
Gets or sets the subtitle.

```yaml
Type: String
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TemplateLayout
Gets or sets a specific template layout by name, one-based position, native type or returned metadata.

```yaml
Type: PowerPointLayoutSelector
Parameter Sets: Template
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Gets or sets the slide title.

```yaml
Type: String
Parameter Sets: Content, Template
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Transition
Gets or sets the transition for this slide and its continuations.

```yaml
Type: PowerPointSlideTransition
Parameter Sets: Content, Template
Aliases: None
Possible values: None, Fade, Wipe, BlindsVertical, BlindsHorizontal, CombHorizontal, CombVertical, PushUp, PushDown, PushLeft, PushRight, Cut, Flash, WarpIn, WarpOut, Prism, FerrisLeft, FerrisRight, Morph, Random, CheckerHorizontal, CheckerVertical, CoverLeft, CoverUp, CoverRight, CoverDown, CoverLeftUp, CoverRightUp, CoverLeftDown, CoverRightDown, Dissolve, FadeThroughBlack, UncoverLeft, UncoverUp, UncoverRight, UncoverDown, UncoverLeftUp, UncoverRightUp, UncoverLeftDown, UncoverRightDown, RandomBarsHorizontal, RandomBarsVertical, StripsLeftUp, StripsRightUp, StripsLeftDown, StripsRightDown, WipeUp, WipeRight, WipeDown, BoxOut, BoxIn, SplitHorizontalOut, SplitHorizontalIn, SplitVerticalOut, SplitVerticalIn, Diamond, Plus, Wedge, Newsflash, WheelOneSpoke, WheelTwoSpokes, WheelThreeSpokes, WheelFourSpokes, WheelEightSpokes, Circle, CutThroughBlack

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.PowerPoint.PowerPointPresentation`

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None

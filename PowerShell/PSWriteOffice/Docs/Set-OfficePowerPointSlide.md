---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficePowerPointSlide
## SYNOPSIS
Changes a selected slide's title, template layout, background or transition.

## SYNTAX
### Default (Default)
```powershell
Set-OfficePowerPointSlide [-Slide] <PowerPointSlideSelector> [-Document <PowerPointDocument>] [-Title <string>] [-TemplateLayout <PowerPointLayoutSelector>] [-Master <int>] [-Background <OfficeColor>] [-Transition <PowerPointSlideTransition>] [-PassThru] [<CommonParameters>]
```

### BackgroundImage
```powershell
Set-OfficePowerPointSlide [-Slide] <PowerPointSlideSelector> -BackgroundImage <string> [-Document <PowerPointDocument>] [-Title <string>] [-TemplateLayout <PowerPointLayoutSelector>] [-Master <int>] [-Transition <PowerPointSlideTransition>] [-PassThru] [<CommonParameters>]
```

### ClearBackground
```powershell
Set-OfficePowerPointSlide [-Slide] <PowerPointSlideSelector> -ClearBackground [-Document <PowerPointDocument>] [-Title <string>] [-TemplateLayout <PowerPointLayoutSelector>] [-Master <int>] [-Transition <PowerPointSlideTransition>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Changes a selected slide's title, template layout, background or transition.

## EXAMPLES

### EXAMPLE 1
```powershell
Set-OfficePowerPointSlide 1 -Document $deck -Title 'Quarterly review'
```


### EXAMPLE 2
```powershell
Get-OfficePowerPointSlide -Document $deck | Set-OfficePowerPointSlide -Transition Fade
```


### EXAMPLE 3
```powershell
Set-OfficePowerPointSlide 2 -Document $deck -TemplateLayout 'Title and Content'
```


## PARAMETERS

### -Background
Gets or sets the background color, including theme roles and transparent literals.

```yaml
Type: OfficeColor
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundImage
Gets or sets a literal local background image path.

```yaml
Type: String
Parameter Sets: BackgroundImage
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ClearBackground
Gets or sets whether to remove the slide's background override.

```yaml
Type: SwitchParameter
Parameter Sets: ClearBackground
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Gets or sets the document for numeric slide and template layout selection; omission uses the enclosing document block.

```yaml
Type: PowerPointDocument
Parameter Sets: Default, BackgroundImage, ClearBackground
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Master
Gets or sets a one-based template master position; defaults to the first.

```yaml
Type: Int32
Parameter Sets: Default, BackgroundImage, ClearBackground
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether the original native slide is returned.

```yaml
Type: SwitchParameter
Parameter Sets: Default, BackgroundImage, ClearBackground
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets a native slide or one-based slide position.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Default, BackgroundImage, ClearBackground
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -TemplateLayout
Gets or sets a template layout by literal name, one-based position, native type or returned metadata. Supply Document when selecting a template layout.

```yaml
Type: PowerPointLayoutSelector
Parameter Sets: Default, BackgroundImage, ClearBackground
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Gets or sets the title text; an empty string clears an existing title.

```yaml
Type: String
Parameter Sets: Default, BackgroundImage, ClearBackground
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Transition
Gets or sets the slide transition.

```yaml
Type: PowerPointSlideTransition
Parameter Sets: Default, BackgroundImage, ClearBackground
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

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficePowerPointSlideTransition
## SYNOPSIS
Sets the transition used when advancing to a slide.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficePowerPointSlideTransition -Transition <PowerPointSlideTransition> [-Slide <PowerPointSlide>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Works on an explicit slide object or on the current slide inside the PowerPoint DSL.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Get-OfficePowerPointSlide -Presentation $ppt -Index 0 | Set-OfficePowerPointSlideTransition -Transition Fade
```

Updates the first slide so it uses the Fade transition.

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
Slide to update (optional inside a slide DSL scope).

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

### -Transition
Transition to apply.

```yaml
Type: PowerPointSlideTransition
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: None, Fade, Wipe, BlindsVertical, BlindsHorizontal, CombHorizontal, CombVertical, PushUp, PushDown, PushLeft, PushRight, Cut, Flash, WarpIn, WarpOut, Prism, FerrisLeft, FerrisRight, Morph, Random, CheckerHorizontal, CheckerVertical, CoverLeft, CoverUp, CoverRight, CoverDown, CoverLeftUp, CoverRightUp, CoverLeftDown, CoverRightDown, Dissolve, FadeThroughBlack, UncoverLeft, UncoverUp, UncoverRight, UncoverDown, UncoverLeftUp, UncoverRightUp, UncoverLeftDown, UncoverRightDown, RandomBarsHorizontal, RandomBarsVertical, StripsLeftUp, StripsRightUp, StripsLeftDown, StripsRightDown, WipeUp, WipeRight, WipeDown, BoxOut, BoxIn, SplitHorizontalOut, SplitHorizontalIn, SplitVerticalOut, SplitVerticalIn, Diamond, Plus, Wedge, Newsflash, WheelOneSpoke, WheelTwoSpokes, WheelThreeSpokes, WheelFourSpokes, WheelEightSpokes, Circle, CutThroughBlack

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

- `OfficeIMO.PowerPoint.PowerPointSlide`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePowerPoint
## SYNOPSIS
Creates a themed presentation, returning a live document or saving a completed content block.

## SYNTAX
### Live (Default)
```powershell
New-OfficePowerPoint [[-Path] <string>] [-Force] [-Theme <OfficeThemeInput>] [-Author <string>] [-Date <DateTime>] [-Accent <OfficeColor>] [-Style <PowerPointDeckStyle>] [-Purpose <string>] [-LayoutStrategy <PowerPointAutoLayoutStrategy>] [-Seed <string>] [-SlideSize <PowerPointSlideSizePreset>] [-NoFooter] [-NoSlideNumber] [-ContinuationFormat <string>] [-Template <string>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Content
```powershell
New-OfficePowerPoint [-Path] <string> [-Content] <scriptblock> [-Force] [-Theme <OfficeThemeInput>] [-Author <string>] [-Date <DateTime>] [-Accent <OfficeColor>] [-Style <PowerPointDeckStyle>] [-Purpose <string>] [-LayoutStrategy <PowerPointAutoLayoutStrategy>] [-Seed <string>] [-SlideSize <PowerPointSlideSizePreset>] [-NoFooter] [-NoSlideNumber] [-ContinuationFormat <string>] [-Template <string>] [-Password <securestring>] [-PassThru] [-AllowLoss] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Without a content block no file is written, including on disposal. With a block, publication is atomic and an existing file requires Force. A block failure leaves the destination unchanged.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficePowerPoint -Path ./Status.pptx -Theme Modern -Author 'Operations team' -PassThru {
    PowerPointSlide -Title 'Status update'
}
```


### EXAMPLE 2
```powershell
$deck = New-OfficePowerPoint -Theme Boardroom
PowerPointSlide -Document $deck -Title 'Service review'
Save-OfficePowerPoint -Document $deck -Path ./Review.pptx
Close-OfficePowerPoint -Document $deck
```


### EXAMPLE 3
```powershell
New-OfficePowerPoint -Path ./Architecture.pptx -Theme Dark -Purpose 'technical proposal' -LayoutStrategy Compact {
    PowerPointSlide -Title 'Architecture review'
    PowerPointSlide -Title 'Decisions' { PowerPointList 'Retain the API boundary', 'Track rollout evidence' }
}
```


## PARAMETERS

### -Accent
Brand accent used to derive the theme's accent palette while retaining its fonts and surfaces.

```yaml
Type: OfficeColor
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowLoss
Explicitly accept reductions reported by the content block's native save.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Author
Author shown on title slides and written to the package's creator property.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Authoring block to complete before the document is atomically saved and closed.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ContinuationFormat
Continuation title format, with title, part and total supplied as {0}, {1} and {2}.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Date
Date shown on title slides using the current culture's short date format.

```yaml
Type: DateTime
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Force
Allow the new document to replace an existing destination.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LayoutStrategy
ContentFirst, DesignFirst, Compact or VisualFirst allocation; overrides the recipe's strategy.

```yaml
Type: PowerPointAutoLayoutStrategy
Parameter Sets: Live, Content
Aliases: None
Possible values: ContentFirst, DesignFirst, Compact, VisualFirst

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoFooter
Suppress footer text on every designed slide.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoSlideNumber
Suppress slide numbers on every designed slide.

```yaml
Type: SwitchParameter
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Return the saved FileInfo after a completed content block.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password for the saved file; plaintext input warns, and credentials contribute their password.

```yaml
Type: SecureString
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Optional associated destination for a live document; required when a content block saves a file.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Purpose
Plain-language purpose matched by the native design brief to a built-in recipe.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Seed
Stable deck seed; an individual slide can supply its own seed.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show
Open the saved result in its registered application after the content block succeeds.

```yaml
Type: SwitchParameter
Parameter Sets: Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SlideSize
Slide dimensions; copied templates retain their dimensions unless this parameter is supplied.

```yaml
Type: PowerPointSlideSizePreset
Parameter Sets: Live, Content
Aliases: None
Possible values: Screen4x3, Screen16x9, Screen16x10

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Composition tuning layered on the theme; Auto uses the theme and purpose-matched recipe.

```yaml
Type: PowerPointDeckStyle
Parameter Sets: Live, Content
Aliases: None
Possible values: Auto, FieldProof, EditorialCaseStudy, TechnicalMap, QuietAppendix

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Template
Existing presentation/template copied as starting content. Theme applies visual styling separately.

```yaml
Type: String
Parameter Sets: Live, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Built-in theme, theme object or a Word, Excel or PowerPoint document/template whose visual theme is read. Defaults to Modern.

```yaml
Type: OfficeThemeInput
Parameter Sets: Live, Content
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

- `OfficeIMO.PowerPoint.PowerPointDocument`
- `System.IO.FileInfo`

## RELATED LINKS

- None

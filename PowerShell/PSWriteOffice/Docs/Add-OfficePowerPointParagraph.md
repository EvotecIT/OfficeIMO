---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointParagraph
## SYNOPSIS
Adds a plain or rich paragraph backed by native content before layout.

## SYNTAX
### Text (Default)
```powershell
Add-OfficePowerPointParagraph [-Text] <string> [-Alignment <OfficeAlignment>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-Bold] [-Italic] [-Strikethrough] [-Underline <PowerPointUnderlineStyle>] [-Color <OfficeColor>] [-Highlight <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-PassThru] [-Slide <PowerPointSlideSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePowerPointParagraph [-Content] <scriptblock> [-Alignment <OfficeAlignment>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-PassThru] [-Slide <PowerPointSlideSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Adds a plain or rich paragraph backed by native content before layout.

## EXAMPLES

### EXAMPLE 1
```powershell
PowerPointSlide -Title 'Status' {
    $paragraph = PowerPointParagraph { PowerPointText 'Healthy' -Bold } -PassThru
    $paragraph.AddText(' in all regions')
}
```


## PARAMETERS

### -Alignment
Gets or sets paragraph alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Text, Content
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Gets or sets whether plain text is bold.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Gets or sets the text color or theme role.

```yaml
Type: OfficeColor
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Column
Gets or sets the one-based theme grid column; explicit X overrides its position.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets the block of native styled runs.

```yaml
Type: ScriptBlock
Parameter Sets: Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Gets or sets the font family.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Gets or sets the height; omission uses the native default for fixed placement.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Highlight
Gets or sets the highlight color.

```yaml
Type: OfficeColor
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Gets or sets whether plain text is italic.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to emit the live native paragraph.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Gets or sets the font size; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Slide
Gets or sets a native slide or one-based position; omission uses the current slide block.

```yaml
Type: PowerPointSlideSelector
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -SpacingAfter
Gets or sets space after the paragraph; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
Gets or sets space before the paragraph; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Span
Gets or sets the grid column count; omission with Column uses twelve. Explicit Width overrides it.

```yaml
Type: Int32
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Gets or sets whether plain text has a strike through it.

```yaml
Type: SwitchParameter
Parameter Sets: Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Gets or sets the paragraph text.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Gets or sets the native underline style.

```yaml
Type: PowerPointUnderlineStyle
Parameter Sets: Text
Aliases: None
Possible values: None, Words, Single, Double, Heavy, Dotted, HeavyDotted, Dash, DashHeavy, DashLong, DashLongHeavy, DotDash, DotDashHeavy, DotDotDash, DotDotDashHeavy, Wavy, WavyHeavy, WavyDouble

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Gets or sets the width; omission uses the native default for fixed placement.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -X
Gets or sets the left coordinate; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Y
Gets or sets the top coordinate; numbers mean points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Content
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

- `PSWriteOffice.Foundation.PowerPointSlideSelector`: A live slide or a one-based slide position in a document.

## OUTPUTS

- `OfficeIMO.PowerPoint.PowerPointParagraph`

## RELATED LINKS

- None

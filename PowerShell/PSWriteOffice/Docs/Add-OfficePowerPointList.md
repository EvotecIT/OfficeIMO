---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePowerPointList
## SYNOPSIS
Adds a native list through the shared layout and continuation rules.

## SYNTAX
### Items (Default)
```powershell
Add-OfficePowerPointList [-Items] <string[]> [-Numbered] [-Numbering <PowerPointNumberingScheme>] [-BulletChar <string>] [-Level <int>] [-PassThru] [-Slide <PowerPointSlideSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [<CommonParameters>]
```

### Content
```powershell
Add-OfficePowerPointList [-Content] <scriptblock> [-Numbered] [-Numbering <PowerPointNumberingScheme>] [-BulletChar <string>] [-Level <int>] [-PassThru] [-Slide <PowerPointSlideSelector>] [-X <OfficeLength>] [-Y <OfficeLength>] [-Width <OfficeLength>] [-Height <OfficeLength>] [-Column <Int32>] [-Span <Int32>] [<CommonParameters>]
```

## DESCRIPTION
Adds a native list through the shared layout and continuation rules.

## EXAMPLES

### EXAMPLE 1
```powershell
$list = PowerPointList 'Availability', 'Capacity' -PassThru
$list.Paragraphs[1].Text = 'Capacity planning'
```


## PARAMETERS

### -BulletChar
Gets or sets a single-character bullet; cannot be combined with numbering.

```yaml
Type: String
Parameter Sets: Items, Content
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
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Content
Gets or sets a block of plain or rich PowerPointListItem commands.

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

### -Height
Gets or sets the height; omission uses the native default for fixed placement.

```yaml
Type: OfficeLength
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Items
Gets or sets the list items in display order.

```yaml
Type: String[]
Parameter Sets: Items
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Level
Gets or sets the nesting level, one or two.

```yaml
Type: Int32
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numbered
Gets or sets whether the list is numbered.

```yaml
Type: SwitchParameter
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numbering
Gets or sets the top-level numbering scheme; supplying it enables numbering.

```yaml
Type: PowerPointNumberingScheme
Parameter Sets: Items, Content
Aliases: None
Possible values: AlphaLowerCharacterParenBoth, AlphaUpperCharacterParenBoth, AlphaLowerCharacterParenR, AlphaUpperCharacterParenR, AlphaLowerCharacterPeriod, AlphaUpperCharacterPeriod, ArabicParenBoth, ArabicParenR, ArabicPeriod, ArabicPlain, RomanLowerCharacterParenBoth, RomanUpperCharacterParenBoth, RomanLowerCharacterParenR, RomanUpperCharacterParenR, RomanLowerCharacterPeriod, RomanUpperCharacterPeriod, CircleNumberDoubleBytePlain, CircleNumberWingdingsBlackPlain, CircleNumberWingdingsWhitePlain, ArabicDoubleBytePeriod, ArabicDoubleBytePlain, EastAsianSimplifiedChinesePeriod, EastAsianSimplifiedChinesePlain, EastAsianTraditionalChinesePeriod, EastAsianTraditionalChinesePlain, EastAsianJapaneseDoubleBytePeriod, EastAsianJapaneseKoreanPlain, EastAsianJapaneseKoreanPeriod, Arabic1Minus, Arabic2Minus, Hebrew2Minus, ThaiAlphaPeriod, ThaiAlphaParenthesisRight, ThaiAlphaParenthesisBoth, ThaiNumberPeriod, ThaiNumberParenthesisRight, ThaiNumberParenthesisBoth, HindiAlphaPeriod, HindiNumPeriod, HindiNumberParenthesisRight, HindiAlpha1Period

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Gets or sets whether to emit the native frame backing the first physical list segment.

```yaml
Type: SwitchParameter
Parameter Sets: Items, Content
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
Parameter Sets: Items, Content
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Span
Gets or sets the grid column count; omission with Column uses twelve. Explicit Width overrides it.

```yaml
Type: Int32
Parameter Sets: Items, Content
Aliases: None
Possible values:

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
Parameter Sets: Items, Content
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
Parameter Sets: Items, Content
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
Parameter Sets: Items, Content
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

- `OfficeIMO.PowerPoint.PowerPointTextBox`

## RELATED LINKS

- None

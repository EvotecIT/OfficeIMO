---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeWordText
## SYNOPSIS
Adds styled text or a hyperlink to the current or given Word paragraph.

## SYNTAX
### Text (Default)
```powershell
Add-OfficeWordText [-Text] <string[]> [-Paragraph <WordParagraph>] [-Bold] [-Italic] [-Strikethrough] [-Underline <WordUnderlineStyle>] [-Color <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-PassThru] [<CommonParameters>]
```

### Link
```powershell
Add-OfficeWordText [-Text] <string[]> -Link <uri> [-Paragraph <WordParagraph>] [-Bold] [-Italic] [-Strikethrough] [-Underline <WordUnderlineStyle>] [-Color <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Tooltip <string>] [-NoHistory] [-NoLinkStyle] [-PassThru] [<CommonParameters>]
```

### Anchor
```powershell
Add-OfficeWordText [-Text] <string[]> -Anchor <string> [-Paragraph <WordParagraph>] [-Bold] [-Italic] [-Strikethrough] [-Underline <WordUnderlineStyle>] [-Color <OfficeColor>] [-Font <string>] [-Size <OfficeLength>] [-Tooltip <string>] [-NoHistory] [-NoLinkStyle] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
A bare Size means points; unit-suffixed lengths are accepted. Hyperlinks use the Hyperlink style unless NoLinkStyle is supplied.

## EXAMPLES

### EXAMPLE 1
```powershell
WordParagraph { WordText 'Important: ' -Bold -Font Aptos -Size 11pt; WordText 'read first' -Italic }
```


### EXAMPLE 2
```powershell
WordParagraph { WordText 'Example' -Link 'https://example.org' -Tooltip 'External reference'; WordText 'Summary' -Anchor Summary }
```


## PARAMETERS

### -Anchor
Bookmark target in this document.

```yaml
Type: String
Parameter Sets: Anchor
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold
Apply bold formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Color
Text color, including a semantic theme role.

```yaml
Type: OfficeColor
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Font family.

```yaml
Type: String
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic
Apply italic formatting.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Link
Absolute hyperlink URL.

```yaml
Type: Uri
Parameter Sets: Link
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHistory
Disable hyperlink navigation history.

```yaml
Type: SwitchParameter
Parameter Sets: Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLinkStyle
Do not apply the built-in Hyperlink character style.

```yaml
Type: SwitchParameter
Parameter Sets: Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Paragraph to receive text outside a content block.

```yaml
Type: WordParagraph
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the target paragraph, or each created hyperlink for Link and Anchor.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Size
Font size; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Strikethrough
Apply strikethrough.

```yaml
Type: SwitchParameter
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Text segments to append.

```yaml
Type: String[]
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Tooltip
Hyperlink tooltip.

```yaml
Type: String
Parameter Sets: Link, Anchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline
Underline style.

```yaml
Type: WordUnderlineStyle
Parameter Sets: Text, Link, Anchor
Aliases: None
Possible values: Single, Words, Double, Thick, Dotted, DottedHeavy, Dash, DashedHeavy, DashLong, DashLongHeavy, DotDash, DashDotHeavy, DotDotDash, DashDotDotHeavy, Wave, WavyHeavy, WavyDouble, None

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordParagraph`

## OUTPUTS

- `OfficeIMO.Word.WordParagraph`
- `OfficeIMO.Word.WordHyperLink`

## RELATED LINKS

- None

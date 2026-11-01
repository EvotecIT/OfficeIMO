---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Set-OfficeWordParagraph
## SYNOPSIS
Updates paragraph style, spacing, indentation, and pagination hints.

## SYNTAX
### Default (Default)
```powershell
Set-OfficeWordParagraph [[-Paragraph] <WordParagraph>] [-Style <WordParagraphStyles>] [-StyleId <string>] [-Alignment <OfficeAlignment>] [-CharacterAlignment <WordVerticalCharacterAlignment>] [-IndentationBefore <OfficeLength>] [-IndentationAfter <OfficeLength>] [-IndentationFirstLine <OfficeLength>] [-IndentationHanging <OfficeLength>] [-LineSpacing <OfficeLength>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-LineSpacingRule <WordLineSpacingRule>] [-PageBreakBefore] [-KeepWithNext] [-KeepLinesTogether] [-AvoidWidowAndOrphan] [-TextDirection <WordTextDirection>] [-BiDi] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates paragraph style, spacing, indentation, and pagination hints.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $p = Add-OfficeWordParagraph -Text 'Executive Summary' -PassThru; $p | Set-OfficeWordParagraph -Style Heading1 -KeepWithNext
```

Applies a heading style and keeps it with the next paragraph.

## PARAMETERS

### -Alignment
Paragraph alignment.

```yaml
Type: OfficeAlignment
Parameter Sets: Default
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AvoidWidowAndOrphan
Enable widow and orphan control.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BiDi
Set or clear right-to-left paragraph layout.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CharacterAlignment
Vertical character alignment on each line.

```yaml
Type: WordVerticalCharacterAlignment
Parameter Sets: Default
Aliases: None
Possible values: Top, Center, Baseline, Bottom, Auto

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationAfter
Indentation after the paragraph; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationBefore
Indentation before the paragraph; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationFirstLine
First-line indentation; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationHanging
Hanging indentation; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepLinesTogether
Keep all paragraph lines together.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeepWithNext
Keep this paragraph with the next paragraph.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineSpacing
Line spacing; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineSpacingRule
Line spacing calculation rule.

```yaml
Type: WordLineSpacingRule
Parameter Sets: Default
Aliases: None
Possible values: Auto, Exact, AtLeast

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageBreakBefore
Start the paragraph on a new page.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Paragraph to update.

```yaml
Type: WordParagraph
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the updated paragraph.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfter
Line spacing after the paragraph; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
Line spacing before the paragraph; a bare number means points.

```yaml
Type: OfficeLength
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style
Paragraph style to apply.

```yaml
Type: WordParagraphStyles
Parameter Sets: Default
Aliases: None
Possible values: Normal, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, ListParagraph, Custom, Title, Subtitle

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StyleId
Paragraph style id to apply.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TextDirection
Paragraph text direction.

```yaml
Type: WordTextDirection
Parameter Sets: Default
Aliases: None
Possible values: LeftToRightTopToBottom, LeftToRightTopToBottom2010, TopToBottomRightToLeft, TopToBottomRightToLeft2010, BottomToTopLeftToRight, BottomToTopLeftToRight2010, LeftToRightTopToBottomRotated, LeftToRightTopToBottomRotated2010, TopToBottomRightToLeftRotated, TopToBottomRightToLeftRotated2010, TopToBottomLeftToRightRotated, TopToBottomLeftToRightRotated2010

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

## RELATED LINKS

- None

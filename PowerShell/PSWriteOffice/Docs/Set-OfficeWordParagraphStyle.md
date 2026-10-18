---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordParagraphStyle
## SYNOPSIS
Updates paragraph style, spacing, indentation, and pagination hints.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeWordParagraphStyle [[-Paragraph] <WordParagraph>] [-Style <WordParagraphStyles>] [-StyleId <string>] [-Alignment <WordParagraphAlignment>] [-CharacterAlignment <WordVerticalCharacterAlignment>] [-IndentationBeforePoints <Double>] [-IndentationAfterPoints <Double>] [-IndentationFirstLinePoints <Double>] [-IndentationHangingPoints <Double>] [-LineSpacingPoints <Double>] [-SpacingBeforePoints <Double>] [-SpacingAfterPoints <Double>] [-LineSpacingRule <WordLineSpacingRule>] [-PageBreakBefore <Boolean>] [-KeepWithNext <Boolean>] [-KeepLinesTogether <Boolean>] [-AvoidWidowAndOrphan <Boolean>] [-TextDirection <WordTextDirection>] [-BiDi <Boolean>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates paragraph style, spacing, indentation, and pagination hints.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $p = Add-OfficeWordParagraph -Text 'Executive Summary' -PassThru; $p | Set-OfficeWordParagraphStyle -Style Heading1 -KeepWithNext $true
```

Applies a heading style and keeps it with the next paragraph.

## PARAMETERS

### -Alignment
Paragraph alignment.

```yaml
Type: WordParagraphAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Left, Start, Center, Right, End, Both, MediumKashida, Distribute, NumTab, HighKashida, LowKashida, ThaiDistribute

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AvoidWidowAndOrphan
Enable widow and orphan control.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
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
Type: Boolean
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Top, Center, Baseline, Bottom, Auto

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationAfterPoints
Indentation after the paragraph in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationBeforePoints
Indentation before the paragraph in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationFirstLinePoints
First-line indentation in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IndentationHangingPoints
Hanging indentation in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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
Type: Boolean
Parameter Sets: __AllParameterSets
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
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LineSpacingPoints
Line spacing in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Type: Boolean
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfterPoints
Line spacing after the paragraph in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBeforePoints
Line spacing before the paragraph in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: Normal, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, ListParagraph, Custom

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
Parameter Sets: __AllParameterSets
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
Parameter Sets: __AllParameterSets
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

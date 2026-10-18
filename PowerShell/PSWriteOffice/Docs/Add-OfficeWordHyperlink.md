---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeWordHyperlink
## SYNOPSIS
Adds a hyperlink to the current Word paragraph.

## SYNTAX
### ContextUrl (Default)
```powershell
Add-OfficeWordHyperlink [-Text] <string> -Url <string> [-Styled] [-Tooltip <string>] [-NoHistory] [-PassThru] [<CommonParameters>]
```

### ParagraphUrl
```powershell
Add-OfficeWordHyperlink [-Text] <string> -Paragraph <WordParagraph> -Url <string> [-Styled] [-Tooltip <string>] [-NoHistory] [-PassThru] [<CommonParameters>]
```

### ParagraphAnchor
```powershell
Add-OfficeWordHyperlink [-Text] <string> -Paragraph <WordParagraph> -Anchor <string> [-Styled] [-Tooltip <string>] [-NoHistory] [-PassThru] [<CommonParameters>]
```

### ContextAnchor
```powershell
Add-OfficeWordHyperlink [-Text] <string> -Anchor <string> [-Styled] [-Tooltip <string>] [-NoHistory] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a hyperlink to the current Word paragraph.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Add-OfficeWordParagraph { Add-OfficeWordHyperlink -Text 'Example' -Url 'https://example.org' -Styled }
```

Creates a styled external hyperlink in the active paragraph.

## PARAMETERS

### -Anchor
Bookmark anchor target within the document.

```yaml
Type: String
Parameter Sets: ParagraphAnchor, ContextAnchor
Aliases: Bookmark
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHistory
Do not mark the hyperlink in navigation history.

```yaml
Type: SwitchParameter
Parameter Sets: ContextUrl, ParagraphUrl, ParagraphAnchor, ContextAnchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Paragraph
Paragraph to update outside the DSL context.

```yaml
Type: WordParagraph
Parameter Sets: ParagraphUrl, ParagraphAnchor
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PassThru
Emit the created hyperlink.

```yaml
Type: SwitchParameter
Parameter Sets: ContextUrl, ParagraphUrl, ParagraphAnchor, ContextAnchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Styled
Apply the built-in hyperlink style.

```yaml
Type: SwitchParameter
Parameter Sets: ContextUrl, ParagraphUrl, ParagraphAnchor, ContextAnchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Displayed hyperlink text.

```yaml
Type: String
Parameter Sets: ContextUrl, ParagraphUrl, ParagraphAnchor, ContextAnchor
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Tooltip
Optional hyperlink tooltip.

```yaml
Type: String
Parameter Sets: ContextUrl, ParagraphUrl, ParagraphAnchor, ContextAnchor
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Url
External hyperlink URL.

```yaml
Type: String
Parameter Sets: ContextUrl, ParagraphUrl
Aliases: Uri
Possible values:

Required: True
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

- `OfficeIMO.Word.WordHyperLink`

## RELATED LINKS

- None

---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownTableOfContents
## SYNOPSIS
Adds a Markdown table of contents placeholder.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownTableOfContents [-Title <string>] [-MinLevel <int>] [-MaxLevel <int>] [-Ordered] [-TitleLevel <int>] [-PlaceAtTop] [-ForPreviousHeading] [-ForSection <string>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownTableOfContents -Document <MarkdownDoc> [-Title <string>] [-MinLevel <int>] [-MaxLevel <int>] [-Ordered] [-TitleLevel <int>] [-PlaceAtTop] [-ForPreviousHeading] [-ForSection <string>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown table of contents placeholder.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\Report.md {
    Add-OfficeMarkdownHeading -Level 1 -Text 'Release Report'
    Add-OfficeMarkdownHeading -Level 2 -Text 'Summary'
    Add-OfficeMarkdownParagraph -Text 'Current state and risks.'
    Add-OfficeMarkdownHeading -Level 2 -Text 'Appendix'
    Add-OfficeMarkdownParagraph -Text 'Supporting evidence.'
    Add-OfficeMarkdownTableOfContents -Title 'Contents' -MinLevel 2 -MaxLevel 3 -PlaceAtTop
}
```

Inserts a generated table of contents for headings in the document.

## PARAMETERS

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDoc
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ForPreviousHeading
Scope the TOC to the previous heading.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ForSection
Scope the TOC to the named section heading.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxLevel
Maximum heading depth included in the table of contents.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MinLevel
Minimum heading depth included in the table of contents.

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Ordered
Generate an ordered table of contents list.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated Markdown document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PlaceAtTop
Insert the TOC at the start of the document.

```yaml
Type: SwitchParameter
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Heading text displayed above the generated table of contents.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleLevel
Heading level used for the TOC title.

```yaml
Type: Int32
Parameter Sets: Context, Document
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

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

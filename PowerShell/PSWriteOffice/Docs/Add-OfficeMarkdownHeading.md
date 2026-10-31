---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeMarkdownHeading
## SYNOPSIS
Adds a Markdown heading.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownHeading [[-Level] <int>] [-Text] <string> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownHeading [[-Level] <int>] [-Text] <string> -Document <MarkdownDocument> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown heading.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\Report.md {
    Add-OfficeMarkdownHeading -Level 1 -Text 'Operational Report'
    Add-OfficeMarkdownTableOfContents -Title 'Contents' -MinLevel 2 -MaxLevel 3 -PlaceAtTop
    Add-OfficeMarkdownHeading -Level 2 -Text 'Overview'
    Add-OfficeMarkdownParagraph -Text 'Current operational state.'
}
```

Appends headings that can be included in a generated table of contents.

## PARAMETERS

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Level
Heading level (1-6).

```yaml
Type: Int32
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: False
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the Markdown document after appending the heading.

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

### -Text
Heading text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDocument`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDocument`

## RELATED LINKS

- None

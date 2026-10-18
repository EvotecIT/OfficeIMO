---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownQuote
## SYNOPSIS
Adds a Markdown block quote.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownQuote [-Text] <string> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownQuote [-Text] <string> -Document <MarkdownDoc> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown block quote.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\ExecutiveSummary.md {
    Add-OfficeMarkdownHeading -Level 1 -Text 'Executive summary'
    Add-OfficeMarkdownQuote -Text 'The release is ready when examples, generated docs, and artifact readback all agree.'
    Add-OfficeMarkdownParagraph -Text 'Detailed validation follows.'
}
```

Appends a quote block to the current Markdown document.

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

### -PassThru
Emit the Markdown document after appending the quote.

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
Quote text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
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

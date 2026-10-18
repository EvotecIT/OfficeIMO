---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownHorizontalRule
## SYNOPSIS
Adds a Markdown horizontal rule.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownHorizontalRule [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownHorizontalRule -Document <MarkdownDoc> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown horizontal rule.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\Report.md {
    Add-OfficeMarkdownHeading -Level 2 -Text 'Summary'
    Add-OfficeMarkdownParagraph -Text 'Key decisions and status.'
    Add-OfficeMarkdownHorizontalRule
    Add-OfficeMarkdownHeading -Level 2 -Text 'Appendix'
}
```

Appends a horizontal rule between sections of the current Markdown document.

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
Emit the Markdown document after appending the rule.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

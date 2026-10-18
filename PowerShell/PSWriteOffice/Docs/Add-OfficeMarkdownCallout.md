---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownCallout
## SYNOPSIS
Adds a Markdown callout block.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownCallout [-Kind] <string> [-Title] <string> [-Body] <string> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownCallout [-Kind] <string> [-Title] <string> [-Body] <string> -Document <MarkdownDoc> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown callout block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\ReleaseNotes.md {
    Add-OfficeMarkdownHeading -Level 1 -Text 'Release notes'
    Add-OfficeMarkdownCallout -Kind 'note' -Title 'Validation' -Body 'Artifacts were generated from deterministic example data.'
    Add-OfficeMarkdownCallout -Kind 'warning' -Title 'Manual step' -Body 'Open the workbook in desktop Excel before publishing pivots.'
}
```

Appends callout blocks while composing a Markdown report.

## PARAMETERS

### -Body
Callout body text.

```yaml
Type: String
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

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

### -Kind
Callout kind (e.g. note, tip, warning).

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

### -PassThru
Emit the Markdown document after appending the callout.

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
Callout title.

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

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

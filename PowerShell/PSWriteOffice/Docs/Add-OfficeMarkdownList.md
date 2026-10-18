---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownList
## SYNOPSIS
Adds a Markdown list.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownList [-Items] <string[]> [-Ordered] [-Start <int>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownList [-Items] <string[]> -Document <MarkdownDoc> [-Ordered] [-Start <int>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown list.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\Plan.md {
    Add-OfficeMarkdownHeading -Level 2 -Text 'Priorities'
    Add-OfficeMarkdownList -Items 'PDF examples', 'PowerPoint layout', 'Word report polish'
    Add-OfficeMarkdownHeading -Level 2 -Text 'Release order'
    Add-OfficeMarkdownList -Items 'Validate artifacts', 'Generate docs', 'Publish release' -Ordered -Start 1
}
```

Appends bullet and ordered lists to the current Markdown document.

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

### -Items
List items to add.

```yaml
Type: String[]
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Ordered
Use an ordered list instead of bullets.

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
Emit the Markdown document after appending the list.

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

### -Start
Starting number for ordered lists.

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

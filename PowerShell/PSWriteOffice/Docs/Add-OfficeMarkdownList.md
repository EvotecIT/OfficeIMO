---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficeMarkdownList
## SYNOPSIS
Adds a Markdown list.

## SYNTAX
### Items (Default)
```powershell
Add-OfficeMarkdownList [-Items] <string[]> [-Numbered] [-Start <int>] [-PassThru] [<CommonParameters>]
```

### DocumentItems
```powershell
Add-OfficeMarkdownList [-Items] <string[]> -Document <MarkdownDocument> [-Numbered] [-Start <int>] [-PassThru] [<CommonParameters>]
```

### DocumentContent
```powershell
Add-OfficeMarkdownList [-Content] <scriptblock> -Document <MarkdownDocument> [-Numbered] [-Start <int>] [-PassThru] [<CommonParameters>]
```

### Content
```powershell
Add-OfficeMarkdownList [-Content] <scriptblock> [-Numbered] [-Start <int>] [-PassThru] [<CommonParameters>]
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
    Add-OfficeMarkdownList -Items 'Validate artifacts', 'Generate docs', 'Publish release' -Numbered -Start 1
}
```

Appends bullet and ordered lists to the current Markdown document.

## PARAMETERS

### -Content
List items composed with MarkdownListItem children.

```yaml
Type: ScriptBlock
Parameter Sets: DocumentContent, Content
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Markdown document to update outside the DSL context.

```yaml
Type: MarkdownDocument
Parameter Sets: DocumentItems, DocumentContent
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
Parameter Sets: Items, DocumentItems
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numbered
Use an ordered list instead of bullets.

```yaml
Type: SwitchParameter
Parameter Sets: Items, DocumentItems, DocumentContent, Content
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
Parameter Sets: Items, DocumentItems, DocumentContent, Content
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
Parameter Sets: Items, DocumentItems, DocumentContent, Content
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

- `OfficeIMO.Markdown.MarkdownDocument`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDocument`

## RELATED LINKS

- None

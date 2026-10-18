---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownFrontMatter
## SYNOPSIS
Adds YAML front matter to a Markdown document.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownFrontMatter [-InputObject] <Object> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownFrontMatter [-InputObject] <Object> -Document <MarkdownDoc> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds YAML front matter to a Markdown document.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> MarkdownFrontMatter -InputObject @{ title = 'Weekly Report'; tags = @('ops','summary') }
```

Sets the document header using the supplied key/value pairs.

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

### -InputObject
Front matter data expressed as a hashtable or object.

```yaml
Type: Object
Parameter Sets: Context, Document
Aliases: Data
Possible values:

Required: True
Position: 0
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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## OUTPUTS

- `OfficeIMO.Markdown.MarkdownDoc`

## RELATED LINKS

- None

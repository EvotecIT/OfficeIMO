---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownDefinitionList
## SYNOPSIS
Adds a Markdown definition list.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownDefinitionList [-Definition] <hashtable> [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownDefinitionList [-Definition] <hashtable> -Document <MarkdownDoc> [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown definition list.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> MarkdownDefinitionList -Definition @{ SLA = 'Service level agreement'; SLO = 'Service level objective' }
```

Appends a definition list built from the provided pairs.

## PARAMETERS

### -Definition
Hashtable of term/definition pairs to render.

```yaml
Type: Hashtable
Parameter Sets: Context, Document
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

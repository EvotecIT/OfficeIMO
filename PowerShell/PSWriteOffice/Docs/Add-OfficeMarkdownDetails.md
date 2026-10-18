---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownDetails
## SYNOPSIS
Adds a collapsible Markdown details block.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownDetails [-Summary] <string> [-Content] <scriptblock> [-Open] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownDetails [-Summary] <string> [-Content] <scriptblock> -Document <MarkdownDoc> [-Open] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a collapsible Markdown details block.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> MarkdownDetails -Summary 'Implementation notes' { MarkdownParagraph -Text 'Hidden by default.' }
```

Appends a details/summary block with nested Markdown content.

## PARAMETERS

### -Content
Nested Markdown content rendered inside the details block.

```yaml
Type: ScriptBlock
Parameter Sets: Context, Document
Aliases: None
Possible values:

Required: True
Position: 1
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

### -Open
Render the details block as open by default.

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

### -Summary
Summary text displayed by the details block.

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

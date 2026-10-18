---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Add-OfficeMarkdownImage
## SYNOPSIS
Adds a Markdown image.

## SYNTAX
### Context (Default)
```powershell
Add-OfficeMarkdownImage [-Path] <string> [-Alt <string>] [-Title <string>] [-Width <Double>] [-Height <Double>] [-PassThru] [<CommonParameters>]
```

### Document
```powershell
Add-OfficeMarkdownImage [-Path] <string> -Document <MarkdownDoc> [-Alt <string>] [-Title <string>] [-Width <Double>] [-Height <Double>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Adds a Markdown image.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> New-OfficeMarkdown -Path .\Report.md {
    Add-OfficeMarkdownHeading -Level 2 -Text 'Architecture'
    Add-OfficeMarkdownImage -Path '.\Assets\architecture.png' -Alt 'Architecture diagram' -Title 'Current architecture' -Width 640
}
```

Appends an image block with accessibility text and optional presentation metadata.

## PARAMETERS

### -Alt
Alt text for the image.

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

### -Height
Optional height in pixels.

```yaml
Type: Double
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
Emit the Markdown document after appending the image.

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

### -Path
Image path or URL.

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

### -Title
Optional title for the image.

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

### -Width
Optional width in pixels.

```yaml
Type: Double
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

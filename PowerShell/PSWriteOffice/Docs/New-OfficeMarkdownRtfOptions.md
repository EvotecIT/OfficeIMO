---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeMarkdownRtfOptions
## SYNOPSIS
Creates native MarkdownRtfExportOptions settings for document operations.

## SYNTAX
### Default (Default)
```powershell
New-OfficeMarkdownRtfOptions [-PreserveRawHtmlAsText] [<CommonParameters>]
```

## DESCRIPTION
Creates native MarkdownRtfExportOptions settings for document operations.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeMarkdownRtfOptions
```


## PARAMETERS

### -PreserveRawHtmlAsText
PreserveRawHtmlAsText override; omission preserves the native default.

```yaml
Type: SwitchParameter
Parameter Sets: Default
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

- `None`

## OUTPUTS

- `OfficeIMO.Rtf.Markdown.MarkdownRtfExportOptions`

## RELATED LINKS

- None

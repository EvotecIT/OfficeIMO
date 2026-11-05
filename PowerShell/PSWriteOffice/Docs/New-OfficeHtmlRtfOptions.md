---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeHtmlRtfOptions
## SYNOPSIS
Creates native HTML/CSS import settings for the HTML-to-Word-to-RTF route.

## SYNTAX
### Default (Default)
```powershell
New-OfficeHtmlRtfOptions [-Font <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [<CommonParameters>]
```

## DESCRIPTION
Creates native HTML/CSS import settings for the HTML-to-Word-to-RTF route.

## EXAMPLES

### EXAMPLE 1
```powershell
New-OfficeHtmlRtfOptions -BasePath 'C:\Path'
```


## PARAMETERS

### -BasePath
Base directory or URI used to resolve resources.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Font
Default font family for imported HTML.

```yaml
Type: String
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetContent
Additional stylesheet source text.

```yaml
Type: String[]
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetPath
Literal stylesheet files applied to the imported HTML.

```yaml
Type: String[]
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

- `OfficeIMO.Word.Html.HtmlWordExportOptions`

## RELATED LINKS

- None

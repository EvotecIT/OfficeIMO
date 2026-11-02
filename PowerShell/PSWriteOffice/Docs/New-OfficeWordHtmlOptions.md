---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeWordHtmlOptions
## SYNOPSIS
Creates native Word-to-HTML options for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeWordHtmlOptions [-Font <string>] [-IncludeFontStyles] [-IncludeListStyles] [-IncludeParagraphClasses] [-IncludeRunClasses] [-IncludeDefaultCss] [-UseImagePaths] [-ExcludeFootnotes] [<CommonParameters>]
```

## DESCRIPTION
Creates native Word-to-HTML options for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeWordHtmlOptions -IncludeParagraphClasses -IncludeRunClasses
Convert-OfficeDocument -Path ./Guide.docx -To Html -OutputPath ./Guide.html -Options $options
```


## PARAMETERS

### -ExcludeFootnotes
Omit footnotes from the exported HTML.

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

### -Font
Font family applied to exported text.

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

### -IncludeDefaultCss
Include the native default stylesheet. Omission preserves the engine default.

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

### -IncludeFontStyles
Include run font information as inline styles.

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

### -IncludeListStyles
Include list style information.

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

### -IncludeParagraphClasses
Emit paragraph styles as CSS classes.

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

### -IncludeRunClasses
Emit run character styles as CSS classes.

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

### -UseImagePaths
Reference image file paths instead of embedding image data.

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

- `OfficeIMO.Word.Html.WordHtmlExportOptions`

## RELATED LINKS

- None

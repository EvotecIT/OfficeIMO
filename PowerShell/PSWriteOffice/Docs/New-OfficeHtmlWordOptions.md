---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficeHtmlWordOptions
## SYNOPSIS
Creates native HTML-to-Word options for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficeHtmlWordOptions [-Font <string>] [-BasePath <string>] [-StylesheetPath <string[]>] [-StylesheetContent <string[]>] [-IncludeListStyles] [-ContinueNumbering] [-SupportsHeadingNumbering] [-RenderPreAsTable] [-TableCaptionPosition <WordHtmlTableCaptionPosition>] [-SectionTagHandling <WordHtmlSectionTagHandling>] [<CommonParameters>]
```

## DESCRIPTION
Creates native HTML-to-Word options for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficeHtmlWordOptions -StylesheetPath ./report.css -RenderPreAsTable
Convert-OfficeDocument -Path ./Report.html -To Word -OutputPath ./Report.docx -Options $options
```


## PARAMETERS

### -BasePath
Filesystem base folder for relative resources, without creating it.

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

### -ContinueNumbering
Continue numbering across imported lists.

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
Font family for imported text.

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

### -IncludeListStyles
Import list style information.

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

### -RenderPreAsTable
Render preformatted blocks inside a single-cell table.

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

### -SectionTagHandling
How HTML section tags map into Word.

```yaml
Type: WordHtmlSectionTagHandling
Parameter Sets: Default
Aliases: None
Possible values: WordSection, Block

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -StylesheetContent
Additional CSS stylesheet text.

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
Explicit literal stylesheet files, resolved through the FileSystem provider.

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

### -SupportsHeadingNumbering
Allow heading numbering during import.

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

### -TableCaptionPosition
Position of an imported table caption.

```yaml
Type: WordHtmlTableCaptionPosition
Parameter Sets: Default
Aliases: None
Possible values: Above, Below

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

- `OfficeIMO.Word.Html.WordHtmlImportOptions`

## RELATED LINKS

- None

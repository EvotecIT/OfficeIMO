---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePowerPointHtmlOptions
## SYNOPSIS
Creates native PowerPoint-to-HTML settings for Convert-OfficeDocument.

## SYNTAX
### Default (Default)
```powershell
New-OfficePowerPointHtmlOptions [-Profile <PowerPointHtmlExportProfile>] [-Theme <OfficeThemeInput>] [-Title <string>] [-IncludeHiddenSlides] [-NoNotes] [-NoTables] [-IncludeHiddenShapes] [-NoExtractionProof] [-NoMasterInventory] [-NoHeading] [-NoSlideHeadings] [-NoRenderDiagnostics] [-NoDefaultStyles] [<CommonParameters>]
```

## DESCRIPTION
Creates native PowerPoint-to-HTML settings for Convert-OfficeDocument.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePowerPointHtmlOptions -Theme Modern -NoNotes
Convert-OfficeDocument -Path ./Review.pptx -To Html -OutputPath ./Review.html -Options $options
```


### EXAMPLE 2
```powershell
$options = New-OfficePowerPointHtmlOptions -Profile VisualReview -NoMasterInventory -NoExtractionProof
Convert-OfficeDocument -Path ./Review.pptx -To Html -OutputPath ./Review.visual.html -Options $options
```


## PARAMETERS

### -IncludeHiddenShapes
Include shapes marked hidden.

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

### -IncludeHiddenSlides
Include slides marked hidden.

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

### -NoDefaultStyles
Omit the shared HTML stylesheet.

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

### -NoExtractionProof
Omit slide-aligned extraction proof.

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

### -NoHeading
Omit the visible document heading; the HTML title is retained.

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

### -NoMasterInventory
Omit the master, layout and theme inventory.

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

### -NoNotes
Omit speaker notes.

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

### -NoRenderDiagnostics
Omit visible visual-review render diagnostics.

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

### -NoSlideHeadings
Omit the heading before each slide.

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

### -NoTables
Omit table content.

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

### -Profile
Semantic slide content or positioned visual review.

```yaml
Type: PowerPointHtmlExportProfile
Parameter Sets: Default
Aliases: None
Possible values: SemanticSlides, VisualReview

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Theme
Shared HTML theme: a built-in name, OfficeTheme or Office document template path. Defaults to Modern.

```yaml
Type: OfficeThemeInput
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Optional HTML document title.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.PowerPoint.Html.PowerPointHtmlSaveOptions`

## RELATED LINKS

- None

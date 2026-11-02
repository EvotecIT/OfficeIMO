---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# New-OfficePdfHtmlOptions
## SYNOPSIS
Creates native semantic or positioned PDF-to-HTML export settings.

## SYNTAX
### Default (Default)
```powershell
New-OfficePdfHtmlOptions [-Profile <PdfHtmlProfile>] [-ImageExportMode <PdfHtmlImageExportMode>] [-MaxEmbeddedImageSize <Int64>] [-NoMetadata] [-NoPageContainers] [-NoImagePlaceholders] [-IncludeLinkAnnotations] [-IncludeFormWidgets] [-Fragment] [-DocumentTitleFallback <string>] [<CommonParameters>]
```

## DESCRIPTION
Creates native semantic or positioned PDF-to-HTML export settings.

## EXAMPLES

### EXAMPLE 1
```powershell
$options = New-OfficePdfHtmlOptions -Profile PositionedReview
Convert-OfficeDocument -Path ./Review.pdf -To Html -OutputPath ./Review.html -Options $options
```


### EXAMPLE 2
```powershell
$options = New-OfficePdfHtmlOptions -Fragment -NoMetadata -NoPageContainers
Convert-OfficeDocument -Path ./Report.pdf -To Html -AsString -Options $options
```


## PARAMETERS

### -DocumentTitleFallback
Document title used when the source PDF has no title.

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

### -Fragment
Return an HTML fragment without the complete document shell.

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

### -ImageExportMode
Use embedded images or placeholders for detected PDF images.

```yaml
Type: PdfHtmlImageExportMode
Parameter Sets: Default
Aliases: None
Possible values: PlaceholderOnly, EmbeddedDataUri

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludeFormWidgets
Include inert form-widget projections; positioned review enables them by default.

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

### -IncludeLinkAnnotations
Include inert link annotations; positioned review enables them by default.

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

### -MaxEmbeddedImageSize
Maximum source image bytes embedded in the HTML; zero embeds no nonempty image.

```yaml
Type: Int64
Parameter Sets: Default
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoImagePlaceholders
Exclude readable image projections.

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

### -NoMetadata
Exclude source document metadata.

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

### -NoPageContainers
Omit semantic page wrappers and page-number metadata.

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
Semantic content or positioned, inert review HTML; semantic is the default.

```yaml
Type: PdfHtmlProfile
Parameter Sets: Default
Aliases: None
Possible values: Semantic, PositionedReview

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

- `OfficeIMO.Html.Pdf.PdfToHtmlOptions`

## RELATED LINKS

- None

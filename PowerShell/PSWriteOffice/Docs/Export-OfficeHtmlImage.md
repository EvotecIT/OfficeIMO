---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Export-OfficeHtmlImage
## SYNOPSIS
Exports an HTML render surface as PNG or SVG with structured diagnostics.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeHtmlImage [-Path] <string> [-OutputPath] <string> [-Format <OfficeImageExportFormat>] [-PageIndex <int>] [-DocumentOptions <HtmlConversionDocumentOptions>] [-RenderOptions <HtmlRenderOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Html
```powershell
Export-OfficeHtmlImage [-OutputPath] <string> -Html <string> [-Format <OfficeImageExportFormat>] [-PageIndex <int>] [-DocumentOptions <HtmlConversionDocumentOptions>] [-RenderOptions <HtmlRenderOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeHtmlImage [-OutputPath] <string> -Document <HtmlConversionDocument> [-Format <OfficeImageExportFormat>] [-PageIndex <int>] [-DocumentOptions <HtmlConversionDocumentOptions>] [-RenderOptions <HtmlRenderOptions>] [-PassThru] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports an HTML render surface as PNG or SVG with structured diagnostics.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> Export-OfficeHtmlImage -Path .\Report.html -OutputPath .\Report.png
```

Uses the dependency-free OfficeIMO HTML renderer. Add -PassThru to receive the structured export result.

## PARAMETERS

### -Document
Shared HTML conversion document.

```yaml
Type: HtmlConversionDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -DocumentOptions
Optional HTML parsing and trust settings for path or markup input.

```yaml
Type: HtmlConversionDocumentOptions
Parameter Sets: Path, Html, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
Output image format.

```yaml
Type: OfficeImageExportFormat
Parameter Sets: Path, Html, Document
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp, Bmp, Pbm, Tga, Icon

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Html
HTML markup to render.

```yaml
Type: String
Parameter Sets: Html
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -OutputPath
Destination PNG, JPEG, TIFF, SVG, or WebP path.

```yaml
Type: String
Parameter Sets: Path, Html, Document
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageIndex
Zero-based rendered page index.

```yaml
Type: Int32
Parameter Sets: Path, Html, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the structured image export result.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Html, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path to an HTML file.

```yaml
Type: String
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -RenderOptions
Optional size, pagination, resource, and rendering settings.

```yaml
Type: HtmlRenderOptions
Parameter Sets: Path, Html, Document
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

- `System.String`
- `OfficeIMO.Html.HtmlConversionDocument`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None

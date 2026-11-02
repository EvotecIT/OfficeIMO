---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# ConvertTo-OfficeDocumentImage
## SYNOPSIS
Renders native pages, slides, sheets, ranges or charts to image results without writing files.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeDocumentImage [-Path] <string[]> [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [<CommonParameters>]
```

### LiteralPath
```powershell
ConvertTo-OfficeDocumentImage -LiteralPath <string[]> [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeDocumentImage [-Document] <Object> [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [<CommonParameters>]
```

### Text
```powershell
ConvertTo-OfficeDocumentImage [-Text] <string> -From <string> [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [<CommonParameters>]
```

## DESCRIPTION
Renders native pages, slides, sheets, ranges or charts to image results without writing files.

## EXAMPLES

### EXAMPLE 1
```powershell
ConvertTo-OfficeDocumentImage -Path ./review.pptx -Page 1-3 -Format Svg | Select-Object Name, Width, Height
```


### EXAMPLE 2
```powershell
ConvertTo-OfficeDocumentImage -Text '<h1>Service status</h1>' -From Html
```


## PARAMETERS

### -BackgroundColor
Image background, including Transparent.

```yaml
Type: OfficeColor
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Chart
Chart name within the selected Excel sheets. Cannot be combined with Range.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned native document, including current unsaved content.

```yaml
Type: Object
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Dpi
Target image resolution in dots per inch; omitted values retain native options.

```yaml
Type: Double
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Format
Output image encoding. File destinations infer it when omitted.

```yaml
Type: OfficeImageFormat
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values: Unknown, Png, Jpeg, Gif, Bmp, Tiff, Svg, Emf, Wmf, Icon, Pcx, Webp, Jpeg2000, Jpeg2000Codestream, Avif, JpegXr, PortableMap, Tga

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -From
Markup source format; use Html.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LiteralPath
Literal paths; accepts FileInfo, PSPath and FullName properties.

```yaml
Type: String[]
Parameter Sets: LiteralPath
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -Options
Native image/load options or hashtables, routed to the selected source formats.

```yaml
Type: Object[]
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
One-based caller-ordered pages, slides or diagram pages; defaults to all.

```yaml
Type: OfficePageRange
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password
Password used only when opening encrypted source files.

```yaml
Type: SecureString
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Input paths, with FileSystem wildcard expansion.

```yaml
Type: String[]
Parameter Sets: Path
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Range
A1 range within the selected Excel sheets.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Image scale; omitted values retain native options.

```yaml
Type: Double
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Excel worksheets selected by name, one-based position or native sheet object.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Text
Literal HTML markup.

```yaml
Type: String
Parameter Sets: Text
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -ThrottleLimit
Maximum number of input files processed concurrently.

```yaml
Type: Int32
Parameter Sets: Path, LiteralPath
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

- `System.String[]`
- `System.Object`
- `System.String`

## OUTPUTS

- `OfficeIMO.Drawing.OfficeImageExportResult`

## RELATED LINKS

- None

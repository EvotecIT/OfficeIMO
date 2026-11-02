---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Export-OfficeDocumentImage
## SYNOPSIS
Renders native pages, slides, sheets, ranges or charts and saves each image atomically.

## SYNTAX
### Path (Default)
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputDirectory] <string> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFile
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputPath] <string> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Export-OfficeDocumentImage [-OutputDirectory] <string> -LiteralPath <string[]> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFile
```powershell
Export-OfficeDocumentImage [-OutputPath] <string> -LiteralPath <string[]> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputDirectory] <string> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentFile
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputPath] <string> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Export-OfficeDocumentImage [-Text] <string> [-OutputDirectory] <string> -From <string> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextFile
```powershell
Export-OfficeDocumentImage [-Text] <string> [-OutputPath] <string> -From <string> [-NoClobber] [-Format <OfficeImageFormat>] [-Page <OfficePageRange>] [-Sheet <OfficeSheetSelector[]>] [-Range <string>] [-Chart <string>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Renders native pages, slides, sheets, ranges or charts and saves each image atomically.

## EXAMPLES

### EXAMPLE 1
```powershell
Export-OfficeDocumentImage -Path ./review.pptx -Page 1-3 -Format Svg -OutputDirectory ./slides
```


### EXAMPLE 2
```powershell
Export-OfficeDocumentImage -LiteralPath './budget[final].xlsx' -Sheet Dashboard -Range A1:H30 -OutputPath ./dashboard.png -NoClobber
```


## PARAMETERS

### -BackgroundColor
Image background, including Transparent.

```yaml
Type: OfficeColor
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Document, DocumentFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Text, TextFile
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
Parameter Sets: LiteralPath, LiteralPathFile
Aliases: PSPath, FullName
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByPropertyName)
Accept wildcard characters: False
```

### -NoClobber
Refuse to replace an existing destination image.

```yaml
Type: SwitchParameter
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Options
Native image/load options or hashtables, routed to the selected source formats.

```yaml
Type: Object[]
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputDirectory
Destination folder for the selected image batch.

```yaml
Type: String
Parameter Sets: Path, LiteralPath, Document, Text
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Exact single-image destination; its extension selects the default encoding.

```yaml
Type: String
Parameter Sets: PathFile, LiteralPathFile, DocumentFile, TextFile
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
One-based caller-ordered pages, slides or diagram pages; defaults to all.

```yaml
Type: OfficePageRange
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Path, PathFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Path, PathFile, LiteralPath, LiteralPathFile, Document, DocumentFile, Text, TextFile
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
Parameter Sets: Text, TextFile
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

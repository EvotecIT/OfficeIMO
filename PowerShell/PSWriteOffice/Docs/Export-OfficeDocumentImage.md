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
Export-OfficeDocumentImage [-Path] <string[]> [-OutputDirectory] <string> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathSheet
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputDirectory] <string> -Sheet <OfficeSheetSelector[]> [-NoClobber] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathRange
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputDirectory] <string> -Range <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathChart
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputDirectory] <string> -Chart <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFile
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputPath] <string> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFileSheet
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputPath] <string> -Sheet <OfficeSheetSelector[]> [-NoClobber] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFileRange
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputPath] <string> -Range <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### PathFileChart
```powershell
Export-OfficeDocumentImage [-Path] <string[]> [-OutputPath] <string> -Chart <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPath
```powershell
Export-OfficeDocumentImage [-OutputDirectory] <string> -LiteralPath <string[]> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathSheet
```powershell
Export-OfficeDocumentImage [-OutputDirectory] <string> -LiteralPath <string[]> -Sheet <OfficeSheetSelector[]> [-NoClobber] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathRange
```powershell
Export-OfficeDocumentImage [-OutputDirectory] <string> -LiteralPath <string[]> -Range <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathChart
```powershell
Export-OfficeDocumentImage [-OutputDirectory] <string> -LiteralPath <string[]> -Chart <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-ThrottleLimit <int>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFile
```powershell
Export-OfficeDocumentImage [-OutputPath] <string> -LiteralPath <string[]> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFileSheet
```powershell
Export-OfficeDocumentImage [-OutputPath] <string> -LiteralPath <string[]> -Sheet <OfficeSheetSelector[]> [-NoClobber] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFileRange
```powershell
Export-OfficeDocumentImage [-OutputPath] <string> -LiteralPath <string[]> -Range <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### LiteralPathFileChart
```powershell
Export-OfficeDocumentImage [-OutputPath] <string> -LiteralPath <string[]> -Chart <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputDirectory] <string> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentSheet
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputDirectory] <string> -Sheet <OfficeSheetSelector[]> [-NoClobber] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentRange
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputDirectory] <string> -Range <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentChart
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputDirectory] <string> -Chart <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentFile
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputPath] <string> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentFileSheet
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputPath] <string> -Sheet <OfficeSheetSelector[]> [-NoClobber] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentFileRange
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputPath] <string> -Range <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### DocumentFileChart
```powershell
Export-OfficeDocumentImage [-Document] <Object> [-OutputPath] <string> -Chart <string> [-NoClobber] [-Sheet <OfficeSheetSelector[]>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Text
```powershell
Export-OfficeDocumentImage [-Text] <string> [-OutputDirectory] <string> -From <string> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### TextFile
```powershell
Export-OfficeDocumentImage [-Text] <string> [-OutputPath] <string> -From <string> [-NoClobber] [-Page <OfficePageRange>] [-Format <OfficeImageExportFormat>] [-Dpi <double>] [-Scale <double>] [-BackgroundColor <OfficeColor>] [-Options <Object[]>] [-Password <securestring>] [-WhatIf] [-Confirm] [<CommonParameters>]
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Chart
A chart name in the selected Excel sheets; excludes Range and Page.

```yaml
Type: String
Parameter Sets: PathChart, PathFileChart, LiteralPathChart, LiteralPathFileChart, DocumentChart, DocumentFileChart
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Caller-owned native document, including current unsaved content.

```yaml
Type: Object
Parameter Sets: Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
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
Type: OfficeImageExportFormat
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
Aliases: None
Possible values: Png, Svg, Jpeg, Tiff, Webp, Bmp, Pbm, Tga, Icon

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
Parameter Sets: LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, Document, DocumentSheet, DocumentRange, DocumentChart, Text
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
Parameter Sets: PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, TextFile
Aliases: None
Possible values:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Page
One-based pages, slides or diagram pages. Cannot be combined with Excel selectors.

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
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: True
```

### -Range
An A1 range in the selected Excel sheets; excludes Chart and Page.

```yaml
Type: String
Parameter Sets: PathRange, PathFileRange, LiteralPathRange, LiteralPathFileRange, DocumentRange, DocumentFileRange
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Image scale; omitted values retain native options.

```yaml
Type: Double
Parameter Sets: Path, PathSheet, PathRange, PathChart, PathFile, PathFileSheet, PathFileRange, PathFileChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFile, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, Document, DocumentSheet, DocumentRange, DocumentChart, DocumentFile, DocumentFileSheet, DocumentFileRange, DocumentFileChart, Text, TextFile
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Sheet
Excel worksheets by name, one-based position or native sheet object.

```yaml
Type: OfficeSheetSelector[]
Parameter Sets: PathSheet, PathRange, PathChart, PathFileSheet, PathFileRange, PathFileChart, LiteralPathSheet, LiteralPathRange, LiteralPathChart, LiteralPathFileSheet, LiteralPathFileRange, LiteralPathFileChart, DocumentSheet, DocumentRange, DocumentChart, DocumentFileSheet, DocumentFileRange, DocumentFileChart
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
Parameter Sets: Path, PathSheet, PathRange, PathChart, LiteralPath, LiteralPathSheet, LiteralPathRange, LiteralPathChart
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

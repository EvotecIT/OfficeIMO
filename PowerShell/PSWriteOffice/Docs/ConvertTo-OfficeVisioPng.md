---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# ConvertTo-OfficeVisioPng
## SYNOPSIS
Exports a Visio document page to native dependency-free PNG.

## SYNTAX
### Path (Default)
```powershell
ConvertTo-OfficeVisioPng [-Path] <string> [-OutputPath <string>] [-PageIndex <int>] [-PixelsPerInch <Double>] [-BackgroundColor <string>] [-Transparent] [-NoText] [-FontFilePath <string>] [-FontFaceName <string>] [-FontCollectionIndex <Int32>] [-NoStencilArtwork] [-NoConnectorLabels] [-NoConnectorLabelOverlapResolution] [-Supersampling <Int32>] [-Open] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### Document
```powershell
ConvertTo-OfficeVisioPng -Document <VisioDocument> [-OutputPath <string>] [-PageIndex <int>] [-PixelsPerInch <Double>] [-BackgroundColor <string>] [-Transparent] [-NoText] [-FontFilePath <string>] [-FontFaceName <string>] [-FontCollectionIndex <Int32>] [-NoStencilArtwork] [-NoConnectorLabels] [-NoConnectorLabelOverlapResolution] [-Supersampling <Int32>] [-Open] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION
Exports a Visio document page to native dependency-free PNG.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> OfficeVisio -Path .\ServiceMap.vsdx { VisioRectangle -Text 'API' -X 2 -Y 4 }
            ConvertTo-OfficeVisioPng -Path .\ServiceMap.vsdx -OutputPath .\ServiceMap.png -PixelsPerInch 144
```

Creates a diagram and exports the first page to a PNG preview.

## PARAMETERS

### -BackgroundColor
Background color name or hex value. Use -Transparent for transparent output.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Visio document object.

```yaml
Type: VisioDocument
Parameter Sets: Document
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -FontCollectionIndex
Optional zero-based font collection index.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFaceName
Optional font face name used when selecting from a font collection.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FontFilePath
Optional TrueType/OpenType font file used for text outlines.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoConnectorLabelOverlapResolution
Do not resolve connector label overlaps at export time.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoConnectorLabels
Do not render connector labels.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoStencilArtwork
Do not render built-in stencil artwork.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoText
Do not render shape text.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Open
Open the PNG after saving.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
Aliases: Show
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputPath
Optional output PNG path.

```yaml
Type: String
Parameter Sets: Path, Document
Aliases: OutPath
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PageIndex
Zero-based page index to export.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Visio .vsdx path.

```yaml
Type: String
Parameter Sets: Path
Aliases: FilePath
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -PixelsPerInch
PNG pixels per Visio inch.

```yaml
Type: Double
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Supersampling
Supersampling factor for smoother raster output.

```yaml
Type: Int32
Parameter Sets: Path, Document
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Transparent
Use transparent PNG background.

```yaml
Type: SwitchParameter
Parameter Sets: Path, Document
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
- `OfficeIMO.Visio.VisioDocument`

## OUTPUTS

- `System.Byte[]`
- `System.IO.FileInfo`

## RELATED LINKS

- None

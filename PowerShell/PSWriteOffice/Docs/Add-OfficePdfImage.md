---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/OfficeIMO
schema: 2.0.0
---
# Add-OfficePdfImage
## SYNOPSIS
Adds a raster image or reusable native visual to PDF content.

## SYNTAX
### Path (Default)
```powershell
Add-OfficePdfImage [-Path] <string> [-Alignment <OfficeAlignment>] [-AlternativeText <string>] [-Link <string>] [-LinkContents <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [<CommonParameters>]
```

### DocumentPath
```powershell
Add-OfficePdfImage [-Path] <string> -Document <PdfDocument> [-Alignment <OfficeAlignment>] [-AlternativeText <string>] [-Link <string>] [-LinkContents <string>] [-PassThru] [-Width <OfficeLength>] [-Height <OfficeLength>] [<CommonParameters>]
```

### InputObject
```powershell
Add-OfficePdfImage -InputObject <Object> [-Alignment <OfficeAlignment>] [-AlternativeText <string>] [-Link <string>] [-LinkContents <string>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-SvgPolicy <OfficeVisualSvgPolicy>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixelCount <Int64>] [-Id <string>] [-Title <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [<CommonParameters>]
```

### DocumentInputObject
```powershell
Add-OfficePdfImage -InputObject <Object> -Document <PdfDocument> [-Alignment <OfficeAlignment>] [-AlternativeText <string>] [-Link <string>] [-LinkContents <string>] [-SpacingBefore <OfficeLength>] [-SpacingAfter <OfficeLength>] [-PassThru] [-SvgPolicy <OfficeVisualSvgPolicy>] [-PointsPerPixel <double>] [-MaximumSvgElements <Int32>] [-MaximumSvgViewportDimension <Double>] [-MaximumSvgViewportPixelCount <Int64>] [-Id <string>] [-Title <string>] [-Width <OfficeLength>] [-Height <OfficeLength>] [<CommonParameters>]
```

## DESCRIPTION
Adds a raster image or reusable native visual to PDF content.

## EXAMPLES

### EXAMPLE 1
```powershell
Add-OfficePdfImage -Path 'C:\Path'
```


### EXAMPLE 2
```powershell
Add-OfficePdfImage -InputObject 'Value' -Document 'Value'
```


### EXAMPLE 3
```powershell
Add-OfficePdfImage -Path 'C:\Path' -Document 'Value'
```


## PARAMETERS

### -Alignment
Alignment for this content.

```yaml
Type: OfficeAlignment
Parameter Sets: Path, DocumentPath, InputObject, DocumentInputObject
Aliases: None
Possible values: Left, Center, Right, Justify, Start, End

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AlternativeText
AlternativeText for this content.

```yaml
Type: String
Parameter Sets: Path, DocumentPath, InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Document
Document for this content.

```yaml
Type: PdfDocument
Parameter Sets: DocumentPath, DocumentInputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Optional output height; numbers mean points and unit suffixes are accepted.

```yaml
Type: OfficeLength
Parameter Sets: Path, DocumentPath, InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Id
Id for a reusable visual input.

```yaml
Type: String
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject
InputObject for this content.

```yaml
Type: Object
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: True
Position: named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Link
Link for this content.

```yaml
Type: String
Parameter Sets: Path, DocumentPath, InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -LinkContents
LinkContents for this content.

```yaml
Type: String
Parameter Sets: Path, DocumentPath, InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumSvgElements
MaximumSvgElements for a reusable visual input.

```yaml
Type: Int32
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumSvgViewportDimension
MaximumSvgViewportDimension for a reusable visual input.

```yaml
Type: Double
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaximumSvgViewportPixelCount
MaximumSvgViewportPixelCount for a reusable visual input.

```yaml
Type: Int64
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
PassThru for this content.

```yaml
Type: SwitchParameter
Parameter Sets: DocumentPath, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Path
Path for this content.

```yaml
Type: String
Parameter Sets: Path, DocumentPath
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PointsPerPixel
PointsPerPixel for a reusable visual input.

```yaml
Type: Double
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingAfter
SpacingAfter for this content.

```yaml
Type: OfficeLength
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SpacingBefore
SpacingBefore for this content.

```yaml
Type: OfficeLength
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -SvgPolicy
SvgPolicy for a reusable visual input.

```yaml
Type: OfficeVisualSvgPolicy
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values: PreserveVector, RasterizeWhenNeeded, RequireVector

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Title for a reusable visual input.

```yaml
Type: String
Parameter Sets: InputObject, DocumentInputObject
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Optional output width; numbers mean points and unit suffixes are accepted.

```yaml
Type: OfficeLength
Parameter Sets: Path, DocumentPath, InputObject, DocumentInputObject
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

- `System.Object`

## OUTPUTS

- `OfficeIMO.Pdf.PdfDocument`

## RELATED LINKS

- None

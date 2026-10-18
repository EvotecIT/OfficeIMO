---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# Set-OfficeWordImage
## SYNOPSIS
Updates OfficeIMO Word image sizing, wrapping, crop, and metadata.

## SYNTAX
### __AllParameterSets
```powershell
Set-OfficeWordImage [-Image] <WordImage> [-Width <Double>] [-Height <Double>] [-Wrap <WordImageTextWrapping>] [-CropTop <Int32>] [-CropBottom <Int32>] [-CropLeft <Int32>] [-CropRight <Int32>] [-Rotation <Int32>] [-Opacity <Int32>] [-HorizontalFlip <Boolean>] [-VerticalFlip <Boolean>] [-Title <string>] [-Description <string>] [-Hidden <Boolean>] [-PassThru] [<CommonParameters>]
```

## DESCRIPTION
Updates OfficeIMO Word image sizing, wrapping, crop, and metadata.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $doc = Get-OfficeWord -Path .\Report.docx
$doc |
    Get-OfficeWordImage |
    Set-OfficeWordImage -Width 320 -Wrap Square -Title 'Report image' -Description 'Image used in the report'
$doc | Save-OfficeWord -Path .\Report-Images.docx
```

Gets OfficeIMO image objects from an open document, applies thin image metadata and layout updates, and saves the document.

## PARAMETERS

### -CropBottom
Bottom crop value.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CropLeft
Left crop value.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CropRight
Right crop value.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CropTop
Top crop value.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Description
Image alternate text metadata.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Height
Image height in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Hidden
Whether the image is hidden.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HorizontalFlip
Horizontally flip the image.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Image
Image to update.

```yaml
Type: WordImage
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: True
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Opacity
Fixed opacity percentage.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru
Emit the updated image.

```yaml
Type: SwitchParameter
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Rotation
Image rotation in degrees.

```yaml
Type: Int32
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Title
Image title metadata.

```yaml
Type: String
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -VerticalFlip
Vertically flip the image.

```yaml
Type: Boolean
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Width
Image width in points.

```yaml
Type: Double
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Wrap
Text wrapping mode.

```yaml
Type: WordImageTextWrapping
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: InLineWithText, Square, Tight, Through, TopAndBottom, BehindText, InFrontOfText

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `OfficeIMO.Word.WordImage`

## OUTPUTS

- `OfficeIMO.Word.WordImage`

## RELATED LINKS

- None

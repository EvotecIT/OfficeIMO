---
external help file: PSWriteOffice-help.xml
Module Name: PSWriteOffice
online version: https://github.com/EvotecIT/PSWriteOffice
schema: 2.0.0
---
# New-OfficePdfVisualComparisonOptions
## SYNOPSIS
Creates discoverable rendering and tolerance settings for Compare-OfficePdfVisual.

## SYNTAX
### __AllParameterSets
```powershell
New-OfficePdfVisualComparisonOptions [-Scale <Double>] [-ChannelTolerance <Byte>] [-AllowedDifferenceRatio <Double>] [-Alignment <PdfVisualPageAlignment>] [-BackgroundColor <string>] [-MaxPages <Int32>] [-MaxPixelsPerImage <Int64>] [-MaxTotalPixels <Int64>] [-MaxTotalOutputBytes <Int64>] [<CommonParameters>]
```

## DESCRIPTION
Creates discoverable rendering and tolerance settings for Compare-OfficePdfVisual.

## EXAMPLES

### EXAMPLE 1
```powershell
PS> $options = New-OfficePdfVisualComparisonOptions -ChannelTolerance 2 -AllowedDifferenceRatio 0.001 -MaxPages 50
Compare-OfficePdfVisual -ReferencePath .\Expected.pdf -DifferencePath .\Actual.pdf -Options $options
```


## PARAMETERS

### -Alignment
Page alignment used for differently sized renders.

```yaml
Type: PdfVisualPageAlignment
Parameter Sets: __AllParameterSets
Aliases: None
Possible values: TopLeft, Center

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllowedDifferenceRatio
Maximum differing-pixel ratio treated as equal.

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

### -BackgroundColor
Background color name or hex value.

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

### -ChannelTolerance
Maximum per-channel byte difference treated as equal.

```yaml
Type: Byte
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxPages
Maximum pages compared.

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

### -MaxPixelsPerImage
Maximum pixels accepted per rendered image.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTotalOutputBytes
Maximum total bytes retained for comparison artifacts.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MaxTotalPixels
Maximum pixels accepted across the comparison.

```yaml
Type: Int64
Parameter Sets: __AllParameterSets
Aliases: None
Possible values:

Required: False
Position: named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Scale
Render scale applied before comparison.

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

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

- `None`

## OUTPUTS

- `OfficeIMO.Pdf.PdfVisualComparisonOptions`

## RELATED LINKS

- None
